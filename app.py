from datetime import datetime
import io
from urllib import response
import pandas as pd
import requests
import time
import json
import sys
import numpy as np 
import re 
import os
# from tkinter import Tk, filedialog
import sys
import streamlit as st



class DomitoryAssignment:


    
    Kakao_API_Key = ""
    ODsay_API_Key = ""
    SCHOOL_ADDRESS = "서울시 구로구 경인로 445" 

    def get_kakao_coordinates(self,address, api_key):
        try:
            url = "https://dapi.kakao.com/v2/local/search/address.json"
            headers = {"Authorization": f"KakaoAK {api_key}"}
            response = requests.get(url, headers=headers, params={"query": address})
            response.raise_for_status()
            data = response.json()
            if not data['documents']: print("No documents found"); return None, None
            return data['documents'][0]['x'], data['documents'][0]['y']
        except: return None, None
    def get_odsay_transit_info(self,origin_coords, dest_coords, api_key=ODsay_API_Key):
        try:
            url = "https://api.odsay.com/v1/api/searchPubTransPathT"
            params = {"apiKey": api_key, "SX": origin_coords[0], "SY": origin_coords[1], "EX": dest_coords[0], "EY": dest_coords[1]}
            response = requests.get(url, params=params)
            response.raise_for_status()
            data = response.json()
        # print(data)
            if "error" in data or not data.get('result') or not data.get('result').get('path'): print("Error in ODsay API response"); return None
            return data['result']['path'][0]
        except: return None
        


 
    # def __init__(self,configfile):
        
        
      
        # if configfile is None:
        #     print("\n[중단] 설정 파일 검증을 통과하지 못했습니다.")
        #     print("설정.xlsx 파일을 수정 후 다시 실행해주세요.")
        #     input("엔터 키를 누르면 종료합니다...")
        #     return
        # self.configfile = configfile
        # self.load_config()


    def load_config(self, configfile):

        try:
            df = pd.read_excel(configfile)
            data = dict(zip(df['항목'].astype(str).str.strip(), df['값']))
            self.Kakao_API_Key = str(data.get('카카오키', '')).strip()
            self.ODsay_API_Key = str(data.get('오디세이키', '')).strip()
        except Exception as e:
            print(f"[오류] 설정 파일 로드 실패: {e}")


    # def select_file(self, title="파일 선택", filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*"))):
    #     root = Tk()
    #     root.withdraw()  # Hide the root window
    #     file_path = filedialog.askopenfilename(title=title, filetypes=filetypes)
    #     root.destroy()
    #     return file_path if file_path else None
    
    #숫자만 강제 추출
    def robust_to_numeric(self,series):
        temp_series = series.astype(str).str.extract(r'(\d+)').astype(float)
        return temp_series.fillna(0)

#핵심 키워드 추출
    def parse_preference_key(self,pref_string):
        if pd.isna(pref_string): return None
        key = str(pref_string).replace('<', '').replace('>', '')
        key = key.split(':', 1)[0].split('(', 1)[0].strip()
        return key

    def calculate_score(self,score):
         '''
         30점 만점을 원하시고, 점수구간이 계단식으로 바꾸어달라 요청     -26-01-15 승우선생님
         '''
         try:
            s = float(score)   
             

            if s > 4.5: # 신입생
                if s>=950 and s<=1000: return 30
                elif s>=900 and s<950: return 25
                elif s>=850 and s<900: return 20
                elif s>=800 and s<850: return 15
                elif s>=750 and s<800: return 10
                elif s>=700 and s<750: return 5
                elif s<700 : return 0
                else :
                    print("잘못된 숫자를 입력하였습니다.")
                    return 0.0
            elif s>=0 and s <= 4.5: # 재학생
                if s==4.5 :return 30
                elif s>=4.0 and s<4.5 :return 25
                elif s>=3.5 and s<4.0 :return 20
                elif s>=3.0 and s<3.5 :return 15
                elif s>=2.5 and s<3.0 :return 10
                elif s<2.5 :return 5
                else :
                    print("잘못된 숫자를 입력하였습니다.")
                    return 0.0
            else :
                if s==None:
                    print("값이 없습니다")
                elif s<0:
                    print("음수는 불가능합니다")
                else:
                    print("잘못된 숫자를 입력하였습니다.")
                return 0.0
         except:
            print("숫자가 아닌 값이 입력되었습니다.")
            return 0.0

    def find_flexible_column(self,df_columns, keywords):
        cols_lower = {str(col).lower(): str(col) for col in df_columns}
        for keyword in keywords:
            if keyword.lower() in cols_lower: return cols_lower[keyword.lower()]
        for col_name_original in df_columns:
            col_lower = str(col_name_original).lower()
        for keyword in keywords:
            if keyword.lower() in col_lower: return col_name_original 
        return None
         
    def assign_room(self,rooms):
        print("\n방 정보 파일을 선택하세요.")
        # room_file = rooms
        # df_rooms = pd.read_excel(room_file)
        df_rooms = rooms    
        if not df_rooms.empty:
            print("방 정보 파일이 성공적으로 로드되었습니다.")
        capacity_col = self.find_flexible_column(df_rooms.columns, ['room', '수용', '인원', '정원'])
        room_gender_col = self.find_flexible_column(df_rooms.columns, ['sex', '성별'])
        room_type_col = self.find_flexible_column(df_rooms.columns, ['Type', '유형', '타입'])
        amount_col = self.find_flexible_column(df_rooms.columns, ['amount', '가격', '금액'])
        if not all([capacity_col, room_gender_col, room_type_col, amount_col]):
            raise ValueError("방 정보 파일에 필수 컬럼이 누락되었습니다.")

        df_rooms[capacity_col] = self.robust_to_numeric(df_rooms[capacity_col]) #방 수용인원 높은 순
        df_rooms[room_gender_col] = df_rooms[room_gender_col].str.strip()
        df_rooms[room_type_col] = df_rooms[room_type_col].apply(self.parse_preference_key)
        
        capacity_grouped = df_rooms.groupby([room_gender_col, room_type_col])[capacity_col].sum()
        self.female_capacity_map = capacity_grouped.loc['여자'].to_dict() if '여자' in capacity_grouped.index else {}
        self.male_capacity_map = capacity_grouped.loc['남자'].to_dict() if '남자' in capacity_grouped.index else {}
        
        df_rooms[amount_col] = self.robust_to_numeric(df_rooms[amount_col]) #방금액 높은 순
        self.room_price_map = df_rooms.drop_duplicates(subset=[room_type_col]).set_index(room_type_col)[amount_col].to_dict()
        
        print(f"-> 정원 및 금액 정보 로드 완료. (여:{sum(self.female_capacity_map.values())}, 남:{sum(self.male_capacity_map.values())})")

        
    def assign_students(self,stu):
       
        print("\n학생 정보 파일을 선택하세요.")
        
        if stu is None:
            print("파일이 선택되지 않았습니다.")
            return
        self.df_students = stu
        # self.df_students = pd.read_excel(students_file)
        
        if not self.df_students.empty:
            print("학생 파일이 성공적으로 로드되었습니다.")

        # 출력된 실제 컬럼명 리스트를 기반으로 키워드 보강
        self.id_col = self.find_flexible_column(self.df_students.columns, ['학번(또는 수험번호)(필수)', '학번', 'ID'])
        self.gender_col = self.find_flexible_column(self.df_students.columns, ['성별(필수)', '성별'])
        self.address_col = self.find_flexible_column(self.df_students.columns, ['현재 등본 상 집주소 입력(필수)', '집주소', '주소'])
        
        # 성적 컬럼은 파일마다 다를 수 있으므로 여러 후보 등록
        self.gpa_col = self.find_flexible_column(self.df_students.columns, [
            '직전학기 평균평점 /신입생 입학점수', 
            '직전학기 평균평점 (선택)', 
            '평점', '성적'
        ])
        
        self.priority_col = self.find_flexible_column(self.df_students.columns, ['우선선발', '우선'])
        self.timestamp_col = self.find_flexible_column(self.df_students.columns, ['타임스탬프', 'Timestamp', '일시']) 
        self.lifepattern_col = self.find_flexible_column(self.df_students.columns, ['생활패턴(필수)', '생활패턴'])

        # 지망 컬럼 (파일에 적힌 실제 긴 제목 추가)
        self.first_choice_col = self.find_flexible_column(self.df_students.columns, ['< 1지망 > 기숙사 실 선택(필수)', '1지망']) 
        self.second_choice_col = self.find_flexible_column(self.df_students.columns, ['< 2지망 > 기숙사 실 선택(필수)', '2지망']) 
        self.third_choice_col = self.find_flexible_column(self.df_students.columns, ['< 3지망 > 기숙사 실 선택(필수)', '3지망']) 
        
        self.account_holder_col = self.find_flexible_column(self.df_students.columns, ['예금주(필수, 학생 본인 계좌이어야 함)', '예금주'])

        # 필수 항목 누락 확인
        required_vars = {
            "학번": self.id_col,
            "성별": self.gender_col,
            "주소": self.address_col,
            "성적": self.gpa_col,
            "1지망": self.first_choice_col
        }
        
        missing = [k for k, v in required_vars.items() if v is None]
        if missing:
            print("\n[오류] 다음 항목을 여전히 찾을 수 없습니다.")
            for m in missing:
                print(f"- {m} (찾으려던 키워드 확인 필요)")
            raise ValueError(f"데이터 파일에서 다음 항목을 찾을 수 없습니다: {', '.join(missing)}")

        # 중복 제거 및 시간순 정렬
        if self.timestamp_col:
            self.df_students[self.timestamp_col] = pd.to_datetime(self.df_students[self.timestamp_col], errors='coerce')
            self.df_students.sort_values(by=self.timestamp_col, ascending=True, inplace=True)
            self.df_students.drop_duplicates(subset=[self.id_col], keep='last', inplace=True)


    def distance_calculation(self):
        print(f"-> 총 {len(self.df_students)}명 학생 거리 계산 중...")
        school_coords = self.get_kakao_coordinates(self.SCHOOL_ADDRESS, self.Kakao_API_Key)
        self.score_results = []
        for i, (idx, row) in enumerate(self.df_students.iterrows()):
            print(f"\r   진행률: {i+1}/{len(self.df_students)}", end="")
           
            sid = row[self.id_col]; addr = row[self.address_col]
            
            if pd.isna(addr):
                self.score_results.append([sid, "주소 없음",0] ); continue
            scoords = self.get_kakao_coordinates(addr, self.Kakao_API_Key)
            if not scoords[0]:
                self.score_results.append([sid, "주소 변환 실패", 0]); continue
                
            rdata = self.get_odsay_transit_info(scoords, school_coords, self.ODsay_API_Key)
            # wcbi, *dets, stat = calculate_wcbi_score(rdata)
            # score_results.append([sid, wcbi, stat] + dets)
            if rdata is None:
                self.score_results.append([sid, "경로 탐색 실패", 0]); continue
            info= rdata.get('info', {})
            total_time = info.get('totalTime', 0)
            subpaths = rdata.get('subPath', []) # ODsay API는 대문자 P를 사용하는 경우도 있으니 확인 필요
            if not subpaths:
                # 경로 정보가 없으면 '경로 없음'으로 처리하고 다음 학생으로
                self.score_results.append([sid, "경로 없음", 0])
                continue

            traffic_what_use = subpaths[0].get('trafficType', 0)
            self.score_results.append([sid, total_time, traffic_what_use])
        
            

            time.sleep(0.05) 
        print("\n-> 거리 계산 완료.")
    def plus_cummute_score(self):
    # 1. 먼저 가중치가 적용된 '원시 점수'를 계산해서 리스트에 보관합니다.
        raw_scores = []
        for row in self.score_results:
            time = row[1]
            traffic = row[2] if len(row) > 2 else 0         
        
            if isinstance(time, (int, float)) and not pd.isna(time):
                # 가중치 적용
                weight = 1.0
                if traffic == 7: weight = 7.0#비행기
                elif traffic == 6: weight = 2.2#시외버스
                elif traffic == 4: weight = 2#기차
                          
                raw_scores.append(float(time) * weight)
            else:
                raw_scores.append(0.0)

        # 2. 데이터 중 가장 높은 점수(MAX)를 찾습니다.
        max_raw_score = max(raw_scores) if raw_scores else 1.0
        if max_raw_score == 0: max_raw_score = 1.0 # 0으로 나누기 방지

        # 3. 최댓값을 70점으로 환산하여 최종 리스트를 만듭니다.
        final_calculated_results = []

        for i, row in enumerate(self.score_results):
           
            sid = row[0]
            time_val = row[1]
            traffic_val = row[2]

            if raw_scores[i] > 0:
                final_score = (raw_scores[i] / max_raw_score) * 70
                # [학번, 시간, 교통, 점수] 형태로 새로 리스트 구성
                final_calculated_results.append([sid, time_val, traffic_val, round(final_score, 2)])
            else:
                final_calculated_results.append([sid, time_val, traffic_val, 0.0])
                
        self.score_results = final_calculated_results # 클래스 변수 업데이트
        print("-> 통학 점수 환산 완료.")
    

    def make_Frame(self):
        scols=[self.id_col,   # 엑셀의 '학번(또는 수험번호)(필수)'와 정확히 일치됨
            '통학시간', 
            '교통수단', 
            '통학 점수(70점)'
            ]
        df_scores = pd.DataFrame(self.score_results, columns=scols)
        df_final =pd.merge(
            self.df_students, 
            df_scores, 
            on=self.id_col, # left_on, right_on 대신 on 하나만 써도 됩니다.
            how='left'
        )

        
        df_final['통학 점수(70점)'] = df_final['통학 점수(70점)'].fillna(0)
        df_final['성적 점수(30점)'] = df_final[self.gpa_col].apply(self.calculate_score)  #환산식 수정 260115

        df_final['최종 점수'] = df_final['통학 점수(70점)'] + df_final['성적 점수(30점)']
        df_final['배정결과'] = '불합격(대기)'; df_final['배정방식'] = '-'; df_final['배정된 방'] = '-'
        df_final[self.gender_col] = df_final[self.gender_col].str.strip().map({'여': '여자', '남': '남자'}).fillna(df_final[self.gender_col])
        df_final['1지망_Key'] = df_final[self.first_choice_col].apply(self.parse_preference_key)
        df_final['2지망_Key'] = df_final[self.second_choice_col].apply(self.parse_preference_key)
        df_final['3지망_Key'] = df_final[self.third_choice_col].apply(self.parse_preference_key)

        pri_mask = pd.notna(df_final[self.priority_col]) & (df_final[self.priority_col] != '') & (df_final[self.priority_col] != False)
        # Priority
        for idx in df_final[pri_mask].sort_values(by='최종 점수', ascending=False).index:
            if pd.isna(df_final.loc[idx, '최종 점수']): continue
            std = df_final.loc[idx]; gen = std[self.gender_col]
            cmap = self.female_capacity_map if gen == '여자' else self.male_capacity_map
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
        gen_indices = df_final[~pri_mask & (df_final['배정결과'] == '불합격(대기)') & pd.notna(df_final['최종 점수'])].index
        choice_cols = [('1지망_Key', '1지망 배정'), ('2지망_Key', '2지망 배정'), ('3지망_Key', '3지망 배정')]
        for _, (ck, method) in enumerate(choice_cols):
            if gen_indices.empty: break
            grouped = df_final.loc[gen_indices].groupby(ck)
            next_round = []
            for k, grp_slice in grouped:
                if not k: next_round.extend(grp_slice.index); continue
                grp_sorted = grp_slice.sort_values(by='최종 점수', ascending=False)
                for idx in grp_sorted.index:
                    gen = df_final.loc[idx,self.gender_col]
                    cmap = self.female_capacity_map if gen == '여자' else self.male_capacity_map
                    if cmap.get(k, 0) > 0:
                        df_final.loc[idx, ['배정결과','배정된 방','배정방식']] = ['합격 (일반선발)', k, method]
                        cmap[k] -= 1
                    else: next_round.append(idx)
            gen_indices = pd.Index(next_round)

        # Random
        unassigned = df_final.loc[gen_indices].sort_values(by='최종 점수', ascending=False).index
        for idx in unassigned:
            gen = df_final.loc[idx, self.gender_col]
            cmap = self.female_capacity_map if gen == '여자' else self.male_capacity_map
            done = False
            for r, s in cmap.items():
                if s > 0:
                    df_final.loc[idx, ['배정결과','배정된 방','배정방식']] = ['합격 (일반선발)', r, '임의 배정']
                    cmap[r] -= 1; done = True; break
            if not done: df_final.loc[idx, '배정결과'] = '불합격(T.O부족)'

        # Waitlist
        w_indices = df_final[df_final['배정결과'].str.startswith('불합격')].index
        for idx in w_indices:
            if pd.isna(df_final.loc[idx, '최종 점수']): df_final.loc[idx, '배정방식'] = '채점 불가 (주소오류)'
            else:
                fk = df_final.loc[idx, '1지망_Key']
                val = f'{fk} (예비)' if fk else '지망 없음 (예비)'
                df_final.loc[idx, '배정된 방'] = val
                df_final.loc[idx, '배정방식'] = '예비 순번'

        # output_file = f"기숙사 배정결과_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        # df_final.to_excel(output_file, index=False)
        # print(f"-> 결과 파일 생성 완료: {output_file}")
        self.df_final = df_final

    def make_excel(self):
        self.df_final['금액'] = self.df_final['배정된 방'].map(self.room_price_map).fillna(0).astype(int)

        self.df_final.sort_values(
                by=[self.gender_col, '배정된 방', '최종 점수'],
                ascending=[True, True, False], 
                inplace=True
            )

        out_cols = list(self.df_students.columns) + [
                '배정결과', '배정방식', '배정된 방', '금액', 
                '최종 점수', '통학 점수(70점)', '성적 점수(30점)','생활 패턴',
                '채점_상태', 'T_기본시간(분)'
            ]
            
        final_cols = out_cols
        if self.account_holder_col in final_cols:
            idx = final_cols.index(self.account_holder_col)
            if '금액' in final_cols: final_cols.remove('금액')
            final_cols.insert(idx+1, '금액')
        else:
                if '금액' in final_cols: final_cols.remove('금액')
                final_cols.append('금액')
                
        final_cols = list(dict.fromkeys(final_cols))
        final_cols = [c for c in final_cols if c in self.df_final.columns]
            
        output_name = f'기숙사_배정_결과_{datetime.now().strftime("%Y%m%d_%H%M")}.xlsx'
        self.df_final[final_cols].to_excel(output_name, index=False)

        print(f"\n[완료] '{output_name}' 파일이 생성되었습니다!")
        print("-> 공실 현황:")
        print("   여자:", {k:v for k,v in self.female_capacity_map.items() if v>0})
        print("   남자:", {k:v for k,v in self.male_capacity_map.items() if v>0})
            
            
        input("\n엔터 키를 누르면 종료합니다...")    
    def make_system_form(self,df_final, room_price_map, gender_col, id_col, lifepattern_col):
    # 1. 이름 변환용 매핑 (짧은 이름 -> 긴 이름)
    # 로직 내부에서는 'A형'으로 쓰지만, 출력할 때는 풀네임으로 바꿔줍니다.
        short_to_long = {
        'A형': 'A형(기숙사형 2인호의 2인실)',
        'B형': 'B형(기숙사형 2인호의 1인실)',
        'C형': 'C형(기숙사형 3인호의 1인실)',
        'D형': 'D형(기숙사형 3인호의 2인실)',
        'E형': 'E형(기숙사형 4인호의 2인실)',
        'F형': 'F형(아파트형 1인실(여학생 전용))',
        'G형': 'G형(아파트형 2인실(여학생 전용))'
    }
    
        output_df = pd.DataFrame()
        
        # 2. 시스템 양식에 맞춘 컬럼 매핑
        form_cols = {
            '기숙사 실': '배정된 방',
            '성별': gender_col,
            '학번': id_col,
            '성명': '성명(필수)',
            '학과(필수)': '학과(필수)',
            '본인 핸드폰 번호': '본인 핸드폰 번호(필수)',
            '흡연여부': '흡연여부(필수, 방배정 시 고려함) - 동양미래대학교 기숙사는 금연 시설입니다.',
            '희망하는 룸메이트 기재': '희망하는 룸메이트 기재(선택)(예시 - 20241236, 홍길동)',
            '생활패턴': lifepattern_col,
            '납부금액': '금액'
        }

        for target, source in form_cols.items():
            if target == '기숙사 실':
                # 원본의 'A형' 등을 위에서 정의한 긴 이름으로 변환
                output_df[target] = df_final['배정된 방'].map(short_to_long).fillna(df_final['배정된 방'])
            elif target == '납부금액':
                # 원본의 'A형' 등을 기준으로 가격표에서 금액 조회
                output_df[target] = df_final['배정된 방'].map(room_price_map).fillna(0).astype(int)
            elif source in df_final.columns:
                output_df[target] = df_final[source]
            else:
                output_df[target] = "-"
                
        # 3. 배정결과가 '합격'인 데이터만 추출
        output_df = output_df[df_final['배정결과'].str.contains('합격')].copy()
        return output_df
def __main__():
    
    st.set_page_config(page_title="🏨 기숙사생 산정 프로그램", layout="wide")
    st.title("🏨 기숙사생 산정 프로그램")
    config_file = st.file_uploader("설정 파일 업로드", type=['xlsx'])
    domitory_assignment = DomitoryAssignment()
    domitory_assignment.load_config(config_file)
        # --- 파일 선택 --
        
    # domitory_assignment.assign_room()
    # domitory_assignment.assign_students()
    
    # #거리 계산o
    # domitory_assignment.distance_calculation()
    # domitory_assignment.plus_cummute_score()

    # #출력
    # domitory_assignment.make_Frame()
    # domitory_assignment.make_excel()

    st.subheader("📁 데이터 업로드")
    col1, col2 = st.columns(2)
    with col1:
        room_file = st.file_uploader("방 정보.xlsx", type=['xlsx'])
    with col2:
        student_file = st.file_uploader("학생 정보.xlsx", type=['xlsx'])

    if room_file and student_file:
        if st.button("🚀 거리 계산", use_container_width=True):
            # 기존 실행 순서 그대로 유지
            with st.spinner("방 정보 로드 중..."):
               domitory_assignment.assign_room(pd.read_excel(room_file))
                
            with st.spinner("학생 정보 분석 중..."):
               domitory_assignment.assign_students(pd.read_excel(student_file))
                
            st.info("📍 거리 계산 시작 (API 호출 중...)")
            domitory_assignment.distance_calculation()
            
            # st.info("⚖️ 통학 점수 환산 중...")
            domitory_assignment.plus_cummute_score()
                
            domitory_assignment.make_Frame()
                
                # 결과 출력 (기존 make_excel 대신 웹 화면 표시 및 다운로드)
            # st.success("✅ 완료!")
            # if hasattr(domitory_assignment, 'df_final'):
            #     st.dataframe(domitory_assignment.df_final)
                    
            #         # 엑셀 다운로드 버튼
            #     output = io.BytesIO()
            #     with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            #         domitory_assignment.df_final.to_excel(writer, index=False)
                    
            #     st.download_button(
            #             label="📥 결과 엑셀 파일 다운로드",
            #             data=output.getvalue(),
            #             file_name=f"기숙사_배정_결과_{datetime.now().strftime('%m%d_%H%M')}.xlsx",
            #             mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            #         )

            st.success("✅ 모든 계산 및 배정이 완료되었습니다!")
            
            # 두 가지 탭으로 나누어 보여주면 훨씬 깔끔합니다
            tab1, tab2 = st.tabs(["📄 방배정 입력용 양식", "📊 전체 배정 근거 데이터"])

            with tab1:
                st.subheader("방배정 데이터 입력용")
                
                # 함수 호출하여 입력용 데이터 생성
                output_df = domitory_assignment.make_system_form(
                    domitory_assignment.df_final,
                    domitory_assignment.room_price_map,
                    domitory_assignment.gender_col,
                    domitory_assignment.id_col,
                    domitory_assignment.lifepattern_col
            )
                
                # 데이터프레임 표시
                st.dataframe(output_df)
                
                # 1번 파일 다운로드
                out1 = io.BytesIO()
                with pd.ExcelWriter(out1, engine='xlsxwriter') as writer:
                    output_df.to_excel(writer, index=False)
                
                st.download_button(
                    label="📥 입력용 양식 다운로드",
                    data=out1.getvalue(),
                    file_name="기숙사_시스템_입력용.xlsx",
                    mime="application/vnd.ms-excel"
                )

            with tab2:
                st.subheader("2. 전체 데이터 (점수/순위 포함)")
                st.dataframe(domitory_assignment.df_final)
                
                # 2번 파일 다운로드
                out2 = io.BytesIO()
                domitory_assignment.df_final.to_excel(out2, index=False, engine='xlsxwriter')
                st.download_button("📥 전체 근거 데이터 다운로드", out2.getvalue(), "기숙사_배정_상세결과.xlsx")
    else:
        st.info("파일을 업로드하여 시작하세요.")

if __name__ == "__main__":
    __main__()