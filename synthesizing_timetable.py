from numpy import printoptions
from z3 import *
import excel_test #excelからデータを取得するためのファイルをインポート
import output_timetable #excelにデータを出力するためのファイルをインポート
import time
import sys

#excelからデータを取得する
excel_test.execution_bring_data()

#授業名
lec = excel_test.lec_list
#教室(16)
lec_room = excel_test.lec_room_list
#曜日・時間(25)
week_time = [
    '月-1', '月-2', '月-3', '月-4', '月-5', '火-1', '火-2', '火-3', '火-4', '火-5', '水-1', '水-2', '水-3', '水-4', '水-5', 
    '木-1', '木-2', '木-3', '木-4', '木-5', '金-1', '金-2', '金-3', '金-4', '金-5'
]
#教師の集合(29)
teacher = excel_test.teacher_list
#教員と担当授業
tl = excel_test.assign_teacher_lec
#学年・クラス(例)23：2(学年)3(学科))
gc = excel_test.assign_grade_lec
#学年・クラスの集合
gc_list = [11, 12, 13, 14, 21, 22, 23, 31, 32, 33, 41, 42, 43, 50]
#授業の種類
kinds = excel_test.kinds_lec
#2講連続開講の授業
conti_lec = excel_test.continue_lec_list
#授業の形態
format_lec = excel_test.format_lec
#授業の区分
classifiction_lec = excel_test.classifiction_lec
#授業の規模(1:0~30,2:31~60,3:60~90,4:90~,5:パソコン室(203),6:パソコン室(G201,G202),7:体育館, 11:研究棟(実験),12:研究棟(普通教室))
capasity_lec = excel_test.capasity_lec
#教室の種類を設定(1:0~30,2:31~60,3:60~90,4:90~,5:パソコン室(203),6:パソコン室(G201,G202),7:体育館, 11:研究棟(実験),12:研究棟(普通教室))
capasity_room = excel_test.capasity_lec_room
#授業に割り当てる教室の数
lec_room_num = excel_test.need_room_num

lt = excel_test.lt

#結果保存用の変数
timetable_data = {}


"""
#z3の変数の宣言
"""
#部屋の集合
room = [[Int("room[%s, %s]" % (l, r)) for r in lec_room] for l in lec]

#コマの集合
koma = [[Int("koma[%s, %s]" % (l, aa)) for aa in week_time] for l in lec]


"""
#ソルバ―の準備
"""
s = Solver()


"""
#関数の宣言
"""
#授業が教員によって行われるか判定する
def assign(teacher_id, lec_id):
    list = tl.get(teacher[teacher_id])
    if list is None:
        return False
    elif lec[lec_id] in list:
        return True
    else:
        return False

#授業がどの学年・クラスで行われるか判定
def grade(l):
    grade_list = gc.get(lec[l])
    if type(grade_list) is int:
        list = [grade_list]
        return list
    else:
        return grade_list

#授業の種類（必修・選択必修・選択）を取得
def judge_kinds(l):
    return kinds.get(lec[l])

#2時間行う授業か判定する
def judge_continue(l):
    return lec[l] in conti_lec

#引数aa2がaa1の次のコマかどうかを判定する
def next(aa1, aa2):
    if aa2-aa1==1:
        if aa1==4 or aa1==9 or aa1==14 or aa1==19 or aa1==24:
            return False
        else:
            return True
    else:
        return False

#aa1が2講か判定する（next関数と組み合わせて使う）
def judge_koma2(aa1):
    if aa1==1 or aa1==6 or aa1==11 or aa1==16 or aa1==21: 
        return False
    else: 
        return True

#授業の形態を判定する
def judge_format(l):
    return format_lec.get(lec[l])

#授業の区分を判定する
def judge_classification(l):
    return classifiction_lec.get(lec[l])

#授業の定員の判定
def judge_capasity_lec(l):
    return capasity_lec.get(lec[l])

#教室の収容人数（種類）の判定
def judge_capasity_room(r):
    return capasity_room.get(lec_room[r])

#授業に配置する教室の数の判定
def judge_lec_room_num(l):
    return lec_room_num.get(lec[l])



"""
#制約の追加
"""
#全ての講義を必ず行う
def all_conduct(room, koma, lec, week_time):
    #2年情報システム工学科の授業の教室
    for i in range(len(lec)):
        for j in range(len(lec_room)):
            s.add(Or( room[i][j]==1, room[i][j]==0 ))

    #2年情報システム工学科の授業
    for i in range(len(lec)):
        for j in range(len(week_time)):
            s.add(Or( koma[i][j]==1, koma[i][j]==0 ))
    for i in range(len(lec)): #全部0はダメ(少なくとも1つは1)
        s.add(Or([koma[i][j]==1 for j in range(len(week_time))]))


#同じ教員の授業を同じ時間に重複させない
def not_dup_teacher(lec, teacher, wt, koma):
    for i in range(len(lec)):
        for i2 in range(len(lec)):
            if i!=i2:
                for p in range(len(teacher)):
                    for w in range(len(wt)):
                        if assign(p, i)==True and assign(p, i2)==True:
                                s.add(Not( And( koma[i][w]==1, koma[i2][w]==1 ) ))


#必修科目と同じコマに他の授業が重複しないようにする
def not_dup_koma(lec, wt, koma):
    for i in range(len(lec)):
        for i2 in range(len(lec)):
            if i != i2:
                if len( set(grade(i)) & set(grade(i2)) ) >= 1: #set(a)&set(b)でlist aとlist bの中身を比較し、同じものがあるかを判定する。同じものがある場合その値が{x ,y}で返って来る。その長さを求め、入っていれば1以上になる。
                    if(judge_kinds(i)=='必修'):
                        for w in range(len(wt)):
                            s.add(Not( And( koma[i][w]==1, koma[i2][w]==1 ) ))


#講義が2講連続である場合
def continue_lec(lec, wt, koma):
    for i in range(len(lec)):
        if judge_continue(i)==True: #2時間行う授業か判定する(6授業)
            list = []
            for w in range(len(wt)):
                for w2 in range(len(wt)):
                    if next(w, w2)==True and judge_koma2(w)==True: #引数w2がwの次のコマかどうかを判定する かつ wが2コマ目ではないとき
                        list.append(koma[i][w])
                        list.append(koma[i][w2])
            s.add(Or( [list[j]+list[j+1]==2 for j in range(0,len(list),2)] ))



#オンライン授業の前後に対面授業を配置しない
def place_format(lec, wt, koma):
    for w in range(len(wt)):
        for w2 in range(len(wt)):
            if next(w, w2)==True or next(w2, w)==True:
                for i in range(len(lec)):
                    for i2 in range(len(lec)):
                        if judge_format(i)!='オンライン' and judge_format(i2)!='対面':
                            if grade(i)==grade(i2):
                                s.add(Not( And(koma[i][w]==1, koma[i2][w2]==1) ))


#指定した科目以外の授業と同じコマに配置しない
def place_classification(lec, wt, koma):
    for i in range(len(lec)):
        for i2 in range(len(lec)):
            if i != i2:
                for w in range(len(wt)):
                    if len( set(grade(i)) & set(grade(i2)) ) >= 1:
                        if judge_kinds(i)!='教職' or judge_kinds(i2)!='教職':
                            if (len(set(judge_classification(i)) & set(judge_classification(i2))) < 1): #iとi2の区分が異なるとき、True。
                                s.add(Not( And(koma[i][w]==1, koma[i2][w]==1) ))


#学生の人数に応じた広さの教室を割り当てる
abc =  [x for x in range(1, 13, 1)]
def assign_room(lec, lec_room, room):
    for i in range(len(lec)):
        if judge_format(i) == '対面':
            jcl = judge_capasity_lec(i)#授業のキャパシティ
            for c in jcl:
                s.add(Or([room[i][r]==1 for r in range(len(lec_room)) if ( c<=4 and judge_capasity_room(r)<=4 and c <= judge_capasity_room(r) ) or ( c==judge_capasity_room(r) )]) )
                s.add(And([room[i][r]==0 for r in  range(len(lec_room)) if c<=4 and judge_capasity_room(r)<=4 and c > judge_capasity_room(r)]))

            for c in list(set(abc) - set(jcl)):
                for r in range(len(lec_room)):
                    if c>=5 and c==judge_capasity_room(r):
                        s.add(room[i][r]==0)
            if (set(jcl) & {1,2,3,4}) == set():
                s.add( [room[i][r]==0 for r in range(len(lec_room)) if judge_capasity_room(r)<=4 ] )

        else:
            s.add(And( [room[i][k]==0 for k in range(len(lec_room))] ))

#教職の授業を同じコマに配置しない
def not_dup_tp_lec(lec, wt, koma):
    for i in range(len(lec)):
        for i2 in range(len(lec)):
            if i != i2:
                for w in range(len(wt)):
                    if len( set(grade(i)) & set(grade(i2)) ) >= 1:
                        if judge_kinds(i)=='教職' and judge_kinds(i2)=='教職':
                            s.add(Not(And( koma[i][w]==1, koma[i2][w]==1 )))
                            #s.add(sum(koma[i][w]+koma[i2][w]) <=1)

#1コマに配置できる授業数の上限を求める
def upperlimit_koma(lec, week_time, koma):
    for w in range(len(week_time)):
        for c in gc_list:
            s.add(Sum( [koma[l][w] for l in range(len(lec)) if c in grade(l)] )<=2)#k


#1授業1コマ,2講連続の場合2コマ
def upperlimit_lec(lec, wt, koma):
    for i in range(len(lec)):
        if judge_continue(i)==True:
            s.add(Sum( [koma[i][w] for w in range(len(wt))] ) <=2)
        else:
            s.add(Sum( [koma[i][w] for w in range(len(wt))] ) <=1)


#配置する教室の数を制限する
def upperlimit_lec_room(lec, lec_room, room):
    for i in range(len(lec)):
        if judge_format(i) == '対面':
            s.add(Sum( [room[i][j] for j in range(len(lec_room))] ) == judge_lec_room_num(i))


#教室の使用時間が重複しないようにする
def not_dup_room_usage(lec, wt, room, koma):
    for i in range(len(lec)):
        for i2 in range(len(lec)):
            if i != i2:
                for w in range(len(wt)):
                    for r in range(len(lec_room)):
                        #s.add(Not( And( room[i][r]==1, koma[i][w]==1, room[i2][r]==1, koma[i2][w]==1 ) ))
                        #s.add((room[i][r]+ koma[i][w]+ room[i2][r]+ koma[i2][w]) <= 3)
                        #s.add(Or( room[i][r]==0, koma[i][w]==0, room[i2][r]==0, koma[i2][w]==0 ))
                        s.add(Or( room[i][r]!=1, koma[i][w]!=1, room[i2][r]!=1, koma[i2][w]!=1 ))
                        #線形計画問題


"""
#制約の入力
"""
print('Start')
start = time.perf_counter()

astart = time.perf_counter()
all_conduct(room, koma, lec, week_time)
print("all_conduct:", time.perf_counter() - astart)

astart = time.perf_counter()
upperlimit_koma(lec, week_time, koma)
print("upperlimit_koma:", time.perf_counter() - astart)

astart = time.perf_counter()
upperlimit_lec(lec, week_time, koma)
print("upperlimit_lec:",time.perf_counter() - astart)

astart = time.perf_counter()
upperlimit_lec_room(lec, lec_room, room)
print("upperlimit_lec_room:",time.perf_counter() - astart)

astart = time.perf_counter()
not_dup_teacher(lec, teacher, week_time, koma)
print("not_dup_teacher:",time.perf_counter() - astart)

astart = time.perf_counter()
not_dup_koma(lec, week_time, koma)
print("not_dup_koma:",time.perf_counter() - astart)

astart = time.perf_counter()
continue_lec(lec, week_time, koma)
print("continue_lec:",time.perf_counter() - astart)

astart = time.perf_counter()
place_format(lec, week_time, koma)
print("place_format:",time.perf_counter() - astart)

astart = time.perf_counter()
place_classification(lec, week_time, koma)
print("place_classification:",time.perf_counter() - astart)

astart = time.perf_counter()
assign_room(lec, lec_room, room)
print("assign_room:",time.perf_counter() - astart)

astart = time.perf_counter()
not_dup_tp_lec(lec, week_time, koma)
print("not_dup_tp_lec:",time.perf_counter() - astart)

astart = time.perf_counter()
not_dup_room_usage(lec, week_time, room, koma)
print("not_dup_room_usage:",time.perf_counter() - astart)
print()


"""
#モデル検査
#モデルの取得・表示
"""
def result2():
    print("モデル検査開始")
    print("モデル検査実行中")
    r = s.check()
    print("検証完了")
    if r == sat:
        m = s.model()

        #整数変数の解を取得
        for i in range(len(lec)):#komaの結果表示
            value = ""
            for j in range(len(week_time)):
                if m[ koma[i][j] ].as_long() == 1:
                    value = lec[i] + ":"
                    for dlt in lt.get(lec[i]):
                        value += ',' + dlt
                    for l in range(len(lec_room)):
                        if m[ room[i][l] ].as_long()==1:
                            value += "\n" + lec_room[l]
                    for lec_grade in grade(i):
                        timetable_data.setdefault(lec_grade, {}).setdefault(week_time[j], []).append(value) #timetable_dataに解を保存する。{学年・クラス:{曜日・時間:[授業名]}, 学年・クラス:{曜日・時間:[授業名, 授業名]}}
    else:
        print(r)


def result3():
    print("モデル検査開始")
    print("モデル検査実行中")
    r = s.check()
    print("検証完了")
    upper = 3
    while r == sat:
        m = s.model()

        #整数変数の解を取得
        for i in range(len(lec)):#komaの結果表示
            value = ""
            for j in range(len(week_time)):
                if m[ koma[i][j] ].as_long() == 1:
                    #print(week_time[j], end="")
                    value = lec[i]+":"
                    for dlt in lt.get(lec[i]):
                        value += ',' + dlt
                    for l in range(len(lec_room)):
                        if m[ room[i][l] ].as_long()==1:
                            value += "\n" + lec_room[l]
                    for lec_grade in grade(i):
                        timetable_data.setdefault(lec_grade, {}).setdefault(week_time[j], []).append(value) #timetable_dataに解を保存する。{学年・クラス:{曜日・時間:[授業名]}, 学年・クラス:{曜日・時間:[授業名, 授業名]}}

        print(timetable_data)

        print()
        print("制約追加中")
        upper = upper-1
        print(upper)
        if upper > 0:
            for w in range(len(week_time)):
                for c in gc_list:
                    s.add(Sum( [koma[l][w] for l in range(len(lec)) if c in grade(l)] ) <= upper)
            
            print("検証中")
            m_time = time.perf_counter()
            r = s.check()
            print("Model Checking Time: ", time.perf_counter() - m_time)
            print("データ削除")
            timetable_data.clear()
        else:
            print("終了")
            break
        
    if r == unsat:
        print(r)


#モデル検査の実行時間
m_time = time.perf_counter()

#モデル検査の実行
result2()
#result3()

print()
print(timetable_data)

print()
print("Model Checking Time: ", time.perf_counter() - m_time)

print()
output_timetable.output_excel(timetable_data, week_time, gc_list) #output_timetableファイルのexcelへの書き込み関数の実行

#プロセス全体の実行時間
process_time = time.perf_counter() - start
print("end Time: ", process_time)

