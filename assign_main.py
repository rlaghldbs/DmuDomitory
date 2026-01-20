from datetime import datetime
from urllib import response
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
        


 
    def __init__(self,configfile="./설정.xlsx"):
        self.configfile = configfile
        self.load_config()
        if configfile is None:
            print("\n[중단] 설정 파일 검증을 통과하지 못했습니다.")
            print("설정.xlsx 파일을 수정 후 다시 실행해주세요.")
            input("엔터 키를 누르면 종료합니다...")
            return


    def load_config(self):

        try:
            df = pd.read_excel(self.configfile)
            data = dict(zip(df['항목'].astype(str).str.strip(), df['값']))
            self.Kakao_API_Key = str(data.get('카카오키', '')).strip()
            self.ODsay_API_Key = str(data.get('오디세이키', '')).strip()
        except Exception as e:
            print(f"[오류] 설정 파일 로드 실패: {e}")


    def select_file(self, title="파일 선택", filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*"))):
        root = Tk()
        root.withdraw()  # Hide the root window
        file_path = filedialog.askopenfilename(title=title, filetypes=filetypes)
        root.destroy()
        return file_path if file_path else None
    
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
         
    def assign_room(self):
        print("\n방 정보 파일을 선택하세요.")
        room_file = self.select_file(title="방 정보 파일 선택")
        df_rooms = pd.read_excel(room_file)
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

        
    def assign_students(self):
       
        print("\n학생 정보 파일을 선택하세요.")
        students_file = self.select_file(title="학생 정보 파일 선택")
        if not students_file:
            print("파일이 선택되지 않았습니다.")
            return

        self.df_students = pd.read_excel(students_file)
        
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
                if traffic == 7: weight = 3.0#비행기
                elif traffic == 6: weight = 1.5#시외버스
            
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
            '통학 점수(70점)']
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
                '최종 점수', '통학 점수(70점)', '성적 점수(30점)',
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




def __main__():
    
    domitory_assignment = DomitoryAssignment("./설정.xlsx")
        # --- 파일 선택 --
    domitory_assignment.assign_room()
    domitory_assignment.assign_students()
    
    #거리 계산o
    domitory_assignment.distance_calculation()
    domitory_assignment.plus_cummute_score()

    #출력
    domitory_assignment.make_Frame()
    domitory_assignment.make_excel()

if __name__ == "__main__":
    __main__()