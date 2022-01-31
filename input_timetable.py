import pandas as pd
from glob import glob

#資源の変数(timetable.pyの資源と対応)
test_list = []

lec_list = [] #○
lec_room_list = [] #○
teacher_list = [] #○
assign_teacher_lec = {} #○
assign_grade_lec = {} #○
kinds_lec = {} #(必修・選択必修・選択) ○
continue_lec_list = [] #○
classifiction_lec = {} #○
format_lec = {} #(対面・オンライン・オンデマンド) ○
capasity_lec = {} #○
capasity_lec_room = {} #○
need_room_num = {}
lt = {}


#excelファイルの取得
filepaths = glob('Excelファイルの場所')
filepath = filepaths[シート番号]
#dataflameでexcelファイル（入力用, 1番目のファイルを取得）を取得
_input_data = pd.read_excel(filepath, sheet_name=ファイルの番号)


#excelファイルのデータの変数
_lec_name_data = _input_data.iloc[: ,0].dropna() #授業名
_assign_teacher_data = _input_data.iloc[:, 1].dropna() #担当教員
_def_grade = _input_data.iloc[:, 2].dropna() #授業の開講学年
_lec_grade_data = _def_grade.astype(int)
_lec_class_data = _input_data.iloc[:, 3].dropna() #授業の開講クラス

_def_continue_time_lec = _input_data.iloc[:, 4].dropna() #講義時間
_continue_time_lec_data = _def_continue_time_lec.astype(int)

_kind_lec_data = _input_data.iloc[:, 5].dropna() #授業の形式
_format_lec_data = _input_data.iloc[:, 6].dropna() #授業の形態
_classification_lec_data = _input_data.iloc[:, 7].dropna() #授業の区分

_capasity_lec_data = _input_data.iloc[:, 8].dropna() #授業のキャパシティ

_def_need_lec_room = _input_data.iloc[:, 9].dropna() #授業に必要な教室の数
_need_lec_room_data = _def_need_lec_room.astype(int)

_affiliation_teacher_data = _input_data.iloc[:, 10].dropna() #所属教員
_lec_room_data = _input_data.iloc[:, 11].dropna() #使用可能教室名

_def_capasity_lec_room = _input_data.iloc[:, 12].dropna() #教室のキャパシティ
capasity_lec_room_data = _def_capasity_lec_room.astype(int)


"""
#関数
"""

#excelファイルから授業名を取得し、リストに追加
def input_lecture(_df_lec):
    for co in _df_lec:
        if (co in lec_list) == False:
            lec_list.append(co)

#excelファイルから教室名を取得し、リストに追加
def input_lecroom(_df_lecroom):
    for ro in _df_lecroom:
        lec_room_list.append(ro)

    if len(lec_room_list) != len(set(lec_room_list)):
        value = list(set(lec_room_list))
        lec_room_list.clear()
        lec_room_list.extend(value) #keyから値を取り出して、set型へ変換することで重複を無くす。

#excelファイルから所属教師名を取得し、リストに追加
def input_teacher(_affiliation_teacher_data):
    for te in _affiliation_teacher_data:
        teacher_list.append(te)

    
    if len(teacher_list) != len(set(teacher_list)):
        value = list(set(teacher_list))
        teacher_list.clear()
        teacher_list.extend(value) #keyから値を取り出して、set型へ変換することで重複を無くす。
    

#excelから授業の種類(必修・選択必修・選択)を取得し、授業とのdictを作成した後、リストに追加
def input_kind_lec(_lec_name_data, _kind_lec_data):
    for ln, kl in zip(_lec_name_data, _kind_lec_data):
        key = ln
        kinds_lec[key] = kl

#excelから授業の形態(対面・オンライン・オンデマンド)を取得し、授業とのdictを作成した後、リストに追加
def input_format_lec(_format_lec_data, _lec_name_data):
    for dl, df in zip(_lec_name_data, _format_lec_data):
        key = dl
        format_lec[key] = df


#授業の開講クラスの分割処理
def split_class_lec(lc):
    _list = []
    list = []

    if type(lc) == int:
        lc = str(lc)
    
    if ',' in lc:
        _list = lc.split(",")
        for li in _list:
            list.append(li.strip())
    else:
        list.append(lc)
        
    return list
#['1,2,3'] → [1, 2, 3] → [21, 22, 23] → {情報処理プロジェクト:[21, 22, 23]}


#学年・クラスの統合
def join_grade_class(_lec_grade_data, _lec_class_data):
    _def_grade_class = []
    a = []
    
    for lc in split_class_lec(_lec_class_data):
        #_def_grade_class.append(int(str(_lec_grade_data) + str(lc)))
        a.append(str(_lec_grade_data) + lc)
        
    #_def_grade_class.append(_list)
    for i in a:
        _def_grade_class.append(int(i))
        
    return _def_grade_class


#授業と学年・クラスを統合
#1.クラスの分割
#2.学年・クラスの統合
#3.授業と学年・クラスでdictを作成
def input_lec_gc(_lec_name_data, _lec_grade_data, _lec_class_data):
    _def_grade_class = []
    
    #学年・クラスの取得
    for lgd, lcd in zip(_lec_grade_data, _lec_class_data):
        _def_grade_class.append(join_grade_class(lgd, lcd))

    #_def_grade_class = str(_def_grade_class)

    for lc, gc in zip(_lec_name_data, _def_grade_class):
        key = lc
        for i in gc:
            #dictのkeyが空なら、デフォルトで空のリスト[]を作成
            assign_grade_lec.setdefault(key, []).append(i)

            if len( assign_grade_lec[key]) != len(set( assign_grade_lec[key])):
                value = list( set( assign_grade_lec[key]) ) #keyから値を取り出して、set型へ変換することで重複を無くす。
                del  assign_grade_lec[key] #指定したkeyの要素を削除
                for v in value:
                     assign_grade_lec.setdefault(key, []).append(v) #再び重複を無くした要素でkeyに追加

    return assign_grade_lec


#excelから連続で行う講義を取得し、リストに追加
def input_continue_lec(_lec_name_data, _continue_time_lec_data):
    for ln, ctl in zip(_lec_name_data, _continue_time_lec_data):
        if ctl==2:
            continue_lec_list.append(ln)

        """
        if len(continue_lec_list) != len(set(continue_lec_list)):
            value = list(set(continue_lec_list))
            continue_lec_list.clear()
            continue_lec_list.extend(value) #keyから値を取り出して、set型へ変換することで重複を無くす。
        """


#excelファイルから授業のキャパシティを取得し、リストに追加
def input_capasity_lec(_lec_name_data, _capasity_lec_data):
    for ln, _def_capasity_lec in zip(_lec_name_data, _capasity_lec_data):
        key = ln
        for cl in split_class_lec(_def_capasity_lec):
            #cl = str(cl).strip
            capasity_lec.setdefault(key, []).append(int(cl)) #int型に変換して格納。split_class_lecでは、分割のために文字列にしているため。
            if len(capasity_lec[key]) != len(set(capasity_lec[key])):
                value = list( set(capasity_lec[key]) )
                del capasity_lec[key]
                #capasity_lec[key] == value
                for v in value:
                    capasity_lec.setdefault(key, []).append(v)



#excelファイルから教室のキャパシティを取得し、リストに追加
def input_capasity_lec_room(_lec_room_data, capasity_lec_room_data):
    for lr, clr in zip(_lec_room_data, capasity_lec_room_data):
        key = lr
        capasity_lec_room[key] = clr


#文字列のの分割処理
def split_string(st):
    list = []
    
    if ',' in st:
        _list = st.split(",")
        for li in _list:
            list.append(li.strip()) #strip()は空白削除。excelからデータを持ってきたとき、先頭or後ろに空白があり、エラーが出たため。
    else:
        list.append(st)
        
    return list

#授業と担当教員を取得し、担当教員がkeyのdictに授業を追加
def input_addication_lec(_lec_name_data, _assign_teacher_data):
    #教員を分ける
    for lc, atd in zip(_lec_name_data, _assign_teacher_data):
        _def_teacher = split_string(atd)
        for key in _def_teacher:
            if key != "なし":
                assign_teacher_lec.setdefault(key, []).append(lc)
            
                #数値計算概論等、学年が違うが同じコマで開講する授業があるため、学年を変えて同一の授業名が入力される。
                #このとき、同じ授業が複数入力されてしまう。
                #それを避けるために、同一の文字列が入力されているとき、それを排除する。
                if len(assign_teacher_lec[key]) != len(set(assign_teacher_lec[key])):
                    value = list( set(assign_teacher_lec[key]) ) #keyから値を取り出して、set型へ変換することで重複を無くす。
                    del assign_teacher_lec[key] #指定したkeyの要素を削除
                    #capasity_lec[key] == value
                    for v in value:
                        assign_teacher_lec.setdefault(key, []).append(v) #再び重複を無くした要素でkeyに追加


#授業に対して、欲しい教室の数の取得し、dictに追加
def input_need_room_num(_lec_name_data, _need_lec_room_data):
    for lc, nlr in zip(_lec_name_data, _need_lec_room_data):
        key = lc
        #need_room_num.setdefault(key, []).append(nlr)
        need_room_num[key] = nlr


#excelから授業の区分を取得し、授業とのdictを作成した後、リストに追加
def input_classification_lec(_lec_name_data, _classification_lec_data):
    for ln, cld in zip(_lec_name_data, _classification_lec_data):
        key = ln
        def_cl = split_string(cld)
        for cl in def_cl:
            classifiction_lec.setdefault(key, []).append(cl)

            """
            if len(classifiction_lec[key]) != len(set(classifiction_lec[key])):
                value = list( set(classifiction_lec[key]) ) #keyから値を取り出して、set型へ変換することで重複を無くす。
                del classifiction_lec[key] #指定したkeyの要素を削除
                for v in value:
                    classifiction_lec.setdefault(key, []).append(v) #再び重複を無くした要素でkeyに追加
            """
def input_lt(_lec_name_data, _assign_teacher_data):
    for lnd, atd in zip(_lec_name_data, _assign_teacher_data):
        key = lnd
        _def_teacher = split_string(atd)
        for at in _def_teacher:
            lt.setdefault(key,[]).append(at)




"""
#関数の宣言
"""
def execution_bring_data():
    #get_test()
    input_lecture(_lec_name_data)
    input_lecroom(_lec_room_data)
    input_teacher(_affiliation_teacher_data)
    input_lec_gc(_lec_name_data, _lec_grade_data, _lec_class_data)
    input_kind_lec(_lec_name_data, _kind_lec_data)
    input_format_lec(_format_lec_data, _lec_name_data)
    input_continue_lec(_lec_name_data, _continue_time_lec_data)
    input_classification_lec(_lec_name_data, _classification_lec_data)
    input_capasity_lec(_lec_name_data, _capasity_lec_data)
    input_capasity_lec_room(_lec_room_data, capasity_lec_room_data)
    input_addication_lec(_lec_name_data, _assign_teacher_data)
    input_need_room_num(_lec_name_data, _need_lec_room_data)
    input_lt(_lec_name_data, _assign_teacher_data)
