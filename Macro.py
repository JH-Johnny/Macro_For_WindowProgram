"""
pip install pywinauto
Window os 한정 실행 program에 Access할 수 있게 해주는 라이브러리이다.

pip install pyautogui
마우스와 키보드를 제어할 수 있게 해주는 라이브러리이다.
키보드 입력은 한글에 대한 적용이 안되어 있어서 pyperclip 모듈로 복사 후 입력해주어야한다. (hotkey('ctrl', 'v'))로

pyautogui.position() - 마우스 좌표 출력
.moveTo(, , 1) - 좌표로 이동
.dragTO(, , 3, button='right') - button을 클릭한 채 3초에 걸쳐 좌표로 이동, default-left
.click(button='right') - 클릭, default-left
.click(x=,y=) - 좌표로 가서 클릭
.click(clicks=2, interval=0.2) - 더블클릭, interval 클릭 간격시간(초)
.scroll(n) - 마우스 스크롤
.write("", interval=0.1) - 키보드 입력, 각 문자마다 입력 간격시간(초)
.keyDown("") - ""키를 누른 상태를 유지
.press("") - 키를 누른다
.KeyUp("") - ""키를 누른 상태를 해제
.screenshot(path) - 스크린샷 찍고 경로에 저장, region=(start_X, start_Y, end_X, end_Y)로 범위 지정 가능
"""
import pyautogui as gui
from pywinauto import application as App
import pandas as pd
import time

# used program name = 'GeocodingTool64.exe'
# used program download url = http://www.biz-gis.com/index.php?mid=pds&document_srl=187250
app = App.Application()
win_program = input("프로그램 경로와 이름은 무엇입니까? (C:\\Users\\gonet\\Desktop\\???.exe)\n>>>")
# C:\Users\??\Desktop\GeocodingTool64_20220811\GeocodingTool64.exe
app.start(win_program)

# Program Information Get
dialog = app.window()

# Setting 1
dialog['button1'].click()
dialog['button3'].click()

# Program coordinate get
def get_coordinate(tree_name):
    dialog = app.window()
    coordinate = dialog[tree_name].wrapper_object().rectangle() # Left, Top, Right, Bottom
    # Plus 10 to each value for Error adjustment
    return (coordinate.left+5, coordinate.top+10)

# Check btn clicked
def btn_check_state(tree_name):
    dialog = app.window()
    # if checked return '1', or '0'
    return dialog[tree_name].wrapper_object().get_check_state()

# delete input data
def delete_contents(tree_name):
    coord = get_coordinate(tree_name)
    gui.click(coord)
    time.sleep(0.1)
    gui.keyDown("Ctrl")
    gui.press("a")
    gui.press("delete")
    time.sleep(0.1)
    gui.keyUp("Ctrl")
    gui.click(coord)

# input data
def input_data(data, tree_name):
    delete_contents(tree_name)
    dialog[tree_name].type_keys(data, with_spaces=True)

# Code Run
if btn_check_state("button3") == 0:
    coordinate = get_coordinate("button3")
    # Mouse Move_Click
    gui.click(coordinate)

# Data read
df = pd.read_excel(r"C:\Users\??\Desktop\??.xlsx")
data_x = df["소재(지번)"].to_list()

# Dict information for Tree_name = (button3, 경위도 버튼) (Edit2, 입력주소 입력상자) (Button7, 한 건씩 처리 버튼)
# (Edit27, 좌표계 검색 결과)
# Macro Code
box = []
for i in data_x[:3]:
    input_data(i, "Edit2")
    gui.click(get_coordinate("Button7"))
    searched_coordinate = dialog.Edit27.wrapper_object().window_text()
    searched_coordinate = searched_coordinate.split("\t")
    searched_coordinate = tuple(map(float, searched_coordinate))
    box.append(searched_coordinate)

result_value = pd.DataFrame(box, columns=["경도", "위도"])
result_value.to_excel(r"C:\Users\??\Desktop\위도_경도_추출.xlsx", index=False)

# Program exit
app.kill(soft=True)