"""
pip install pywinauto
Window os 한정 실행 program에 Access할 수 있게 해주는 라이브러리이다.
"""
from pywinauto import application as App
import pandas as pd

# used program name = 'GeocodingTool64.exe'
# used program download url = http://www.biz-gis.com/index.php?mid=pds&document_srl=187250
app = App.Application()
win_program = input("프로그램 경로와 이름은 무엇입니까? (C:\\Users\\??\\Desktop\\???.exe)\n>>>")
# C:\Users\??\Desktop\GeocodingTool64_20220811\GeocodingTool64.exe
app.start(win_program)

# Program Information Get
dialog = app.window()

# Setting 1
dialog['button1'].click()
dialog['button3'].click()

# Check btn clicked
def btn_check_state(tree_name):
    dialog = app.window()
    # if checked return '1', or '0'
    return dialog[tree_name].wrapper_object().get_check_state()

# input data
def input_data(data, tree_name):
    # Delete text
    dialog.Edit2.wrapper_object().set_edit_text("")
    # Input text
    dialog[tree_name].type_keys(data, with_spaces=True)

# Code Run
if btn_check_state("button3") == 0:
    dialog.button3.wrapper_object().click()

# Data read
df = pd.read_excel(r"C:\Users\??\Desktop\03. 서천군 음식점정보_v2.0.xlsx")
data_x = df["주소(지번)"].to_list()

# Dict information for Tree_name = (button3, 경위도 버튼) (Edit2, 입력주소 입력상자) (Button7, 한 건씩 처리 버튼)
# (Edit27, 좌표계 검색 결과) (Edit30, 우편번호 결과값)
# Macro Code
box = []
for i in data_x:
    input_data(i, "Edit2")
    dialog.button7.wrapper_object().click()
    searched_coordinate = dialog.Edit30.wrapper_object().window_text()
    #searched_coordinate = searched_coordinate.split("\t")
    try:
        searched_coordinate = int(searched_coordinate)
    except(ValueError) as q:
        print(q, ">>> 우편번호 정보가 없습니다. 0으로 대체합니다.")
        searched_coordinate = 0
    print(searched_coordinate)
    box.append(searched_coordinate)

result_value = pd.DataFrame(box, columns=["우편번호"])
result_value.to_excel(r"C:\Users\??\Desktop\우편번호_추출.xlsx", index=False)

# Program exit
app.kill(soft=True)