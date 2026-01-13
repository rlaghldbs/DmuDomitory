import pandas as pd
import requests
import time
import json
import sys
import numpy as np 
import re 
import os
from tkinter import Tk, filedialog
import sys
sys.stdout.reconfigure(encoding='utf-8')

# --- 1. 기본 설정 및 전역 변수 ---
KAKAO_API_KEY = ""
ODSAY_API_KEY = "IJr37xVwsdbjQHbZjClE9eoFgJoQgFR3G8Sp4p/TA6Y"
SCHOOL_ADDRESS = "서울시 구로구 경인로 445" 

# WCBI 및 배정 파라미터 (설정.xlsx에서 업데이트됨)
MONTHLY_DAYS = 20
F_TRANSFER_PENALTY = 0.15
F_HEADWAY_PENALTY = 0.2
HEADWAY_THRESHOLD = 15
WCBI_WEIGHT = 0.7  
GPA_WEIGHT = 0.3   

# 경제적 부담 계수 동적 리스트 (설정.xlsx 로드 시 생성)
COST_INTERVALS = [] # [(임계값, 계수), ...]
MAX_COST_FACTOR = 2.0

F_MODE_PENALTIES = {
    "AIR": 0.8, "KTX_SRT": 0.5, "EXPRESS_BUS": 0.4,
    "INTERCITY_BUS": 0.4, "TRAIN_NORMAL": 0.3, "WIDE_BUS_1": 0.1,
    "NORMAL": 0.0
}

# --- 2. 엑셀 설정 로더 함수 ---
def load_config_from_excel(config_file='./설정.xlsx'):
    global KAKAO_API_KEY, ODSAY_API_KEY, WCBI_WEIGHT, GPA_WEIGHT
    global HEADWAY_THRESHOLD, COST_INTERVALS, MAX_COST_FACTOR
    
    if not os.path.exists(config_file):
        print(f"[알림] '{config_file}' 파일이 없습니다. 기본 설정을 사용하거나 중단합니다.")
        return None
    else :
        print(f"[알림] '{config_file}' 파일을 로드합니다...")

    try:
        df = pd.read_excel(config_file)
        data = dict(zip(df['항목'].astype(str).str.strip(), df['값']))
        print(data)

        # 1. API 키 로드
        KAKAO_API_KEY = str(data.get('카카오키', '')).strip()
        if( not KAKAO_API_KEY ):
            print("[경고] 카카오 API 키가 설정 파일에 없습니다.")

        ODSAY_API_KEY = str(data.get('오디세이키', '')).strip()
        if( not ODSAY_API_KEY ):
            print("[경고] ODsay API 키가 설정 파일에 없습니다.")

        # 2. 가중치 및 기본 기준 로드
        WCBI_WEIGHT = float(data.get('통학가중치', 0.7))
        GPA_WEIGHT = float(data.get('성적가중치', 0.3))
        HEADWAY_THRESHOLD = int(data.get('배차간격기준', 15))

        # 3. [핵심] 경제적 부담 구간 동적 로드
        # 설정.xlsx에 '비용구간1', '계수1' ... 식으로 정의되어 있다고 가정
        intervals = []
        for i in range(1, 5): # 1~4구간
            threshold = data.get(f'비용구간{i}')
            factor = data.get(f'계수{i}')
            if pd.notna(threshold) and pd.notna(factor):
                intervals.append((int(threshold), float(factor)))
        
        COST_INTERVALS = sorted(intervals, key=lambda x: x[0])
        MAX_COST_FACTOR = float(data.get('계수5', 2.0)) # 4구간 초과 시 적용할 계수

        return data

    except Exception as e:
        print(f"[오류] 설정 파일 로드 중 에러 발생: {e}")
        return None

# --- 헬퍼 함수 ---
def select_file(title):
    root = Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    file_path = filedialog.askopenfilename(title=title, filetypes=[("Excel files", "*.xlsx *.xls")])
    root.destroy()
    return file_path

def find_flexible_column(df_columns, keywords):
    cols_lower = {str(col).lower(): str(col) for col in df_columns}
    for keyword in keywords:
        if keyword.lower() in cols_lower: return cols_lower[keyword.lower()]
    for col_name_original in df_columns:
        col_lower = str(col_name_original).lower()
        for keyword in keywords:
            if keyword.lower() in col_lower: return col_name_original 
    return None

def robust_to_numeric(series):
    temp_series = series.astype(str).str.extract(r'(\d+)').astype(float)
    return temp_series.fillna(0)

def parse_preference_key(pref_string):
    if pd.isna(pref_string): return None
    key = str(pref_string).replace('<', '').replace('>', '')
    key = key.split(':', 1)[0].split('(', 1)[0].strip()
    return key

def get_kakao_coordinates(address, api_key):
    try:
        url = "https://dapi.kakao.com/v2/local/search/address.json"
        headers = {"Authorization": f"KakaoAK {api_key}"}
        response = requests.get(url, headers=headers, params={"query": address})
        response.raise_for_status()
        data = response.json()
        if not data['documents']: print("No documents found"); return None, None
        return data['documents'][0]['x'], data['documents'][0]['y']
    except: return None, None

def get_odsay_transit_info(origin_coords, dest_coords, api_key=ODSAY_API_KEY):
    try:
        url = "https://api.odsay.com/v1/api/searchPubTransPathT"
        params = {"apiKey": api_key, "SX": origin_coords[0], "SY": origin_coords[1], "EX": dest_coords[0], "EY": dest_coords[1]}
        response = requests.get(url, params=params)
        response.raise_for_status()
        data = response.json()
        print(data)
        if "error" in data or not data.get('result') or not data.get('result').get('path'): print("Error in ODsay API response"); return None
        return data['result']['path'][0]
    except: return None
    
# --- 3. 로직 함수 ---

def get_interval_cost_factor(monthly_cost):
    """
    [V9.5.1] 설정.xlsx에서 로드된 구간 정보를 바탕으로 가중치 반환
    """
    for threshold, factor in COST_INTERVALS:
        if monthly_cost <= threshold:
            return factor
    return MAX_COST_FACTOR

def calculate_converted_score_linear(score):
    try:
        s = float(score)
    except: return 0.0
    if s > 10.0: # 신입생
        converted = (s - 650) / 3
    else: # 재학생
        converted = (s - 1.5) * (100 / 3)
    return max(0.0, min(100.0, converted))

def calculate_wcbi_score(route):
    if not route: return 0, 1.0, 1.0, 0.0, 0, 0.0, 0.0, 0.0, 0, "경로 없음"
    info = route['info']; T = info.get('totalTime', 0)
    
    # 월 교통비 계산 및 설정 파일 기준 가중치 적용
    api_payment = info.get('totalPayment', 0)
    calculated_monthly_cost = api_payment * 2 * MONTHLY_DAYS
    Factor_cost = get_interval_cost_factor(calculated_monthly_cost)
    
    # 피로도 계산
    f_transfer = 0.0; f_mode = 0.0; f_headway = 0.0; transit_segment_count = 0
    if 'subPath' in route:
        for sub_path in route['subPath']:
            traffic_type = sub_path.get('trafficType')
            current_penalty = 0.0
            if traffic_type == 7: current_penalty = F_MODE_PENALTIES["AIR"]
            elif traffic_type == 4:
                train_type = str(sub_path.get('trainType', ''))
                current_penalty = F_MODE_PENALTIES["KTX_SRT"] if train_type in ['1', '8'] else F_MODE_PENALTIES["TRAIN_NORMAL"]
            elif traffic_type == 5: current_penalty = F_MODE_PENALTIES["EXPRESS_BUS"]
            elif traffic_type == 6: current_penalty = F_MODE_PENALTIES["INTERCITY_BUS"]
            f_mode = max(f_mode, current_penalty)
            if traffic_type != 3: transit_segment_count += 1
            if traffic_type != 3 and sub_path.get('intervalTime', 0) > HEADWAY_THRESHOLD:
                f_headway = F_HEADWAY_PENALTY
                
    f_transfer = max(0, transit_segment_count - 1) * F_TRANSFER_PENALTY
    Factor_friction = 1 + f_transfer + f_mode + f_headway
    
    # 공식 적용: WCBI = T * 비용계수 * 피로도계수
    WCBI = T * Factor_cost * Factor_friction
    return (WCBI, T, Factor_cost, Factor_friction, max(0, transit_segment_count - 1), f_transfer, f_mode, f_headway, calculated_monthly_cost, "성공")

def main():
    global KAKAO_API_KEY, ODSAY_API_KEY
    global BASELINE_MONTHLY_COST, HEADWAY_THRESHOLD, WCBI_WEIGHT, GPA_WEIGHT
    global F_TRANSFER_PENALTY, MODEL_ALPHA

    print("=== 기숙사 배정 자동화 프로그램 V9.5.2 (설정 파일 연동 강화) ===")
    
    # 1. 엑셀 설정 로드
    config = load_config_from_excel('./설정.xlsx')
    
    if config is None:
        print("\n[중단] 설정 파일 검증을 통과하지 못했습니다.")
        print("설정.xlsx 파일을 수정 후 다시 실행해주세요.")
        input("엔터 키를 누르면 종료합니다...")
        return

    # 2. 검증된 값 할당
    try:
        KAKAO_API_KEY = str(config['카카오키']).strip()
        ODSAY_API_KEY = str(config['오디세이키']).strip()
        
        # 실제 계산에 쓰일 변수 업데이트
        global BASELINE_MONTHLY_COST, HEADWAY_THRESHOLD, WCBI_WEIGHT, GPA_WEIGHT
        HEADWAY_THRESHOLD = int(config.get('배차간격기준', 15))
        WCBI_WEIGHT = float(config.get('통학가중치', 0.7))
        GPA_WEIGHT = float(config.get('성적가중치', 0.3))
        MONTHLY_DAYS = int(config.get('등교일수', 20))
        F_TRANSFER_PENALTY = float(config.get('환승패널티', 0.15))
        
        print(f"-> 설정 검증 완료: 통학 {WCBI_WEIGHT*100}% / 성적 {GPA_WEIGHT*100}% 반영")
        print(f"-> 기준 정보: 등교일수 {MONTHLY_DAYS}일, 환승패널티 {F_TRANSFER_PENALTY}")
        
    except Exception as e:
        print(f"[오류] 데이터 할당 중 예기치 못한 에러: {e}")
        return
    
    # 키가 없으면 사용자 입력 요청
    if not KAKAO_API_KEY: KAKAO_API_KEY = input("카카오 API 키를 입력하세요: ").strip()
    if not ODSAY_API_KEY: ODSAY_API_KEY = input("ODsay API 키를 입력하세요: ").strip()

    if not KAKAO_API_KEY or not ODSAY_API_KEY:
        print("[오류] API 키가 없습니다. 프로그램을 종료합니다.")
        input("엔터 키를 누르면 종료합니다..."); return

    # --- 파일 선택 ---
    print("\n[1단계] '방 정보(정원)' 엑셀 파일을 선택해주세요.")
    room_file = select_file("방 정보 엑셀 파일 선택")
    if not room_file: return

    print("\n[2단계] '학생 신청 데이터' 엑셀 파일을 선택해주세요.")
    student_file = select_file("학생 데이터 엑셀 파일 선택")
    if not student_file: return

    try:
        # Phase 1
        print("\n--- 데이터 분석 및 배정 시작 ---")
        df_rooms = pd.read_excel(room_file)
        capacity_col = find_flexible_column(df_rooms.columns, ['room', '수용', '인원', '정원'])
        room_gender_col = find_flexible_column(df_rooms.columns, ['sex', '성별'])
        room_type_col = find_flexible_column(df_rooms.columns, ['Type', '유형', '타입'])
        amount_col = find_flexible_column(df_rooms.columns, ['amount', '가격', '금액'])
        
        if not all([capacity_col, room_gender_col, room_type_col, amount_col]):
            raise ValueError("방 정보 파일에 필수 컬럼이 누락되었습니다.")

        df_rooms[capacity_col] = robust_to_numeric(df_rooms[capacity_col])
        df_rooms[room_gender_col] = df_rooms[room_gender_col].str.strip()
        df_rooms[room_type_col] = df_rooms[room_type_col].apply(parse_preference_key)
        
        capacity_grouped = df_rooms.groupby([room_gender_col, room_type_col])[capacity_col].sum()
        female_capacity_map = capacity_grouped.loc['여자'].to_dict() if '여자' in capacity_grouped.index else {}
        male_capacity_map = capacity_grouped.loc['남자'].to_dict() if '남자' in capacity_grouped.index else {}
        
        df_rooms[amount_col] = robust_to_numeric(df_rooms[amount_col])
        room_price_map = df_rooms.drop_duplicates(subset=[room_type_col]).set_index(room_type_col)[amount_col].to_dict()
        
        print(f"-> 정원 및 금액 정보 로드 완료. (여:{sum(female_capacity_map.values())}, 남:{sum(male_capacity_map.values())})")

        # Phase 2
        df_students = pd.read_excel(student_file)
        id_col = find_flexible_column(df_students.columns, ['학번'])
        gender_col = find_flexible_column(df_students.columns, ['성별'])
        address_col = find_flexible_column(df_students.columns, ['집주소', '주소'])
        gpa_col = find_flexible_column(df_students.columns, ['직전학기 평균평점', '평점', '학점'])
        priority_col = find_flexible_column(df_students.columns, ['우선 선발 여부', '우선선발', '우선'])
        timestamp_col = find_flexible_column(df_students.columns, ['타임', '일시', 'Timestamp']) 
        first_choice_col = find_flexible_column(df_students.columns, ['1지망']) 
        second_choice_col = find_flexible_column(df_students.columns, ['2지망']) 
        third_choice_col = find_flexible_column(df_students.columns, ['3지망']) 
        account_holder_col = find_flexible_column(df_students.columns, ['예금주'])

        if not all([id_col, gender_col, address_col, gpa_col, priority_col, first_choice_col]):
             raise ValueError("학생 데이터 파일에 필수 컬럼이 누락되었습니다.")

        if timestamp_col:
            df_students[timestamp_col] = pd.to_datetime(df_students[timestamp_col], errors='coerce')
            df_students.sort_values(by=timestamp_col, ascending=True, inplace=True)
            df_students.drop_duplicates(subset=[id_col], keep='last', inplace=True)
        
        # Phase 3
        print(f"-> 총 {len(df_students)}명 학생 거리 계산 중...")
        school_coords = get_kakao_coordinates(SCHOOL_ADDRESS, KAKAO_API_KEY)
        score_results = []
        for i, (idx, row) in enumerate(df_students.iterrows()):
            print(f"\r   진행률: {i+1}/{len(df_students)}", end="")
            sid = row[id_col]; addr = row[address_col]
            if pd.isna(addr):
                score_results.append([sid, np.nan, "주소 없음"] + [np.nan]*8); continue
            scoords = get_kakao_coordinates(addr, KAKAO_API_KEY)
            if not scoords[0]:
                score_results.append([sid, np.nan, "주소 변환 실패"] + [np.nan]*8); continue
            rdata = get_odsay_transit_info(scoords, school_coords, ODSAY_API_KEY)
            wcbi, *dets, stat = calculate_wcbi_score(rdata)
            score_results.append([sid, wcbi, stat] + dets)
            time.sleep(0.05) 
        print("\n-> 거리 계산 완료.")

        scols = [id_col, 'WCBI_최종점수', '채점_상태', 'T_기본시간(분)', '경제적 부담', '통학 피로도', '환승횟수', '환승 페널티', '교통수단 페널티', '배차 페널티', '적용_월교통비']
        df_scores = pd.DataFrame(score_results, columns=scols)
        df_final = pd.merge(df_students, df_scores, on=id_col, how='left')

        # Phase 4 //통학점수 계산
       #원 코드 df_final['통학 점수(100점)'] = df_final['WCBI_최종점수'].rank(method='dense', pct=True) * 100
        
        df_final['통학 점수(100점)'] = df_final['WCBI_최종점수'].rank(method='average', pct=True) * 100
        # 신입생 1000점 / 재학생 4.5점 자동 구분 함수 적용
        df_final['성적 점수(100점)'] = df_final[gpa_col].apply(calculate_converted_score_linear)
        
        df_final['Final_Score (최종 합산점)'] = (df_final['통학 점수(100점)'] * WCBI_WEIGHT) + (df_final['성적 점수(100점)'] * GPA_WEIGHT)

        # Phase 5
        df_final['배정결과'] = '불합격(대기)'; df_final['배정방식'] = '-'; df_final['배정된 방'] = '-'
        df_final[gender_col] = df_final[gender_col].str.strip().map({'여': '여자', '남': '남자'}).fillna(df_final[gender_col])
        df_final['1지망_Key'] = df_final[first_choice_col].apply(parse_preference_key)
        df_final['2지망_Key'] = df_final[second_choice_col].apply(parse_preference_key)
        df_final['3지망_Key'] = df_final[third_choice_col].apply(parse_preference_key)
        
        pri_mask = pd.notna(df_final[priority_col]) & (df_final[priority_col] != '') & (df_final[priority_col] != False)
        
        # Priority
        for idx in df_final[pri_mask].sort_values(by='Final_Score (최종 합산점)', ascending=False).index:
            if pd.isna(df_final.loc[idx, 'Final_Score (최종 합산점)']): continue
            std = df_final.loc[idx]; gen = std[gender_col]
            cmap = female_capacity_map if gen == '여자' else male_capacity_map
            chs = [std['1지망_Key'], std['2지망_Key'], std['3지망_Key']]
            done = False
            for i, c in enumerate(chs):
                if c and cmap.get(c, 0) > 0:
                    df_final.loc[idx, ['배정결과','배정된 방','배정방식']] = ['합격 (우선선발)', c, f'{i+1}지망 배정 (우선)']
                    cmap[c] -= 1; done = True; break
            if not done:
                for r, s in cmap.items():
                    if s > 0:
                        df_final.loc[idx, ['배정결과','배정된 방','배정방식']] = ['합격 (우선선발)', r, '임의 배정 (우선)']
                        cmap[r] -= 1; done = True; break
        
        # General
        gen_indices = df_final[~pri_mask & (df_final['배정결과'] == '불합격(대기)') & pd.notna(df_final['Final_Score (최종 합산점)'])].index
        choice_cols = [('1지망_Key', '1지망 배정'), ('2지망_Key', '2지망 배정'), ('3지망_Key', '3지망 배정')]
        for _, (ck, method) in enumerate(choice_cols):
            if gen_indices.empty: break
            grouped = df_final.loc[gen_indices].groupby(ck)
            next_round = []
            for k, grp_slice in grouped:
                if not k: next_round.extend(grp_slice.index); continue
                grp_sorted = grp_slice.sort_values(by='Final_Score (최종 합산점)', ascending=False)
                for idx in grp_sorted.index:
                    gen = df_final.loc[idx, gender_col]
                    cmap = female_capacity_map if gen == '여자' else male_capacity_map
                    if cmap.get(k, 0) > 0:
                        df_final.loc[idx, ['배정결과','배정된 방','배정방식']] = ['합격 (일반선발)', k, method]
                        cmap[k] -= 1
                    else: next_round.append(idx)
            gen_indices = pd.Index(next_round)

        # Random
        unassigned = df_final.loc[gen_indices].sort_values(by='Final_Score (최종 합산점)', ascending=False).index
        for idx in unassigned:
            gen = df_final.loc[idx, gender_col]
            cmap = female_capacity_map if gen == '여자' else male_capacity_map
            done = False
            for r, s in cmap.items():
                if s > 0:
                    df_final.loc[idx, ['배정결과','배정된 방','배정방식']] = ['합격 (일반선발)', r, '임의 배정']
                    cmap[r] -= 1; done = True; break
            if not done: df_final.loc[idx, '배정결과'] = '불합격(T.O부족)'

        # Waitlist
        w_indices = df_final[df_final['배정결과'].str.startswith('불합격')].index
        for idx in w_indices:
            if pd.isna(df_final.loc[idx, 'Final_Score (최종 합산점)']): df_final.loc[idx, '배정방식'] = '채점 불가 (주소오류)'
            else:
                fk = df_final.loc[idx, '1지망_Key']
                val = f'{fk} (예비)' if fk else '지망 없음 (예비)'
                df_final.loc[idx, '배정된 방'] = val
                df_final.loc[idx, '배정방식'] = '예비 순번'

        # Phase 8: Sorting & Saving
        df_final['금액'] = df_final['배정된 방'].map(room_price_map).fillna(0).astype(int)
        
        df_final.sort_values(
            by=[gender_col, '배정된 방', 'Final_Score (최종 합산점)'],
            ascending=[True, True, False], 
            inplace=True
        )

        out_cols = list(df_students.columns) + [
            '배정결과', '배정방식', '배정된 방', '금액', 
            'Final_Score (최종 합산점)', '통학 점수(100점)', '성적 점수(100점)',
            'WCBI_최종점수', '채점_상태', 'T_기본시간(분)', '경제적 부담', '통학 피로도',
            '적용_월교통비', '환승횟수', '환승 페널티', '교통수단 페널티', '배차 페널티'
        ]
        
        final_cols = out_cols
        if account_holder_col in final_cols:
            idx = final_cols.index(account_holder_col)
            if '금액' in final_cols: final_cols.remove('금액')
            final_cols.insert(idx+1, '금액')
        else:
            if '금액' in final_cols: final_cols.remove('금액')
            final_cols.append('금액')
            
        final_cols = list(dict.fromkeys(final_cols))
        final_cols = [c for c in final_cols if c in df_final.columns]
        
        output_name = '기숙사_배정_결과_최종.xlsx'
        df_final[final_cols].to_excel(output_name, index=False)
        
        print(f"\n[완료] '{output_name}' 파일이 생성되었습니다!")
        print("-> 공실 현황:")
        print("   여자:", {k:v for k,v in female_capacity_map.items() if v>0})
        print("   남자:", {k:v for k,v in male_capacity_map.items() if v>0})
        
        input("\n엔터 키를 누르면 종료합니다...")

    except Exception as e:
        print(f"\n[오류 발생] {e}")
        import traceback
        traceback.print_exc()
        input("엔터 키를 누르면 종료합니다...")

if __name__ == "__main__":
    main()