import xlrd

workbook=xlrd.open_workbook('E:\Python Quiz\Question.xls')

def Question_set(workbook,name,question_no=None):
    worksheet=workbook.sheet_by_name(name)
    if question_no==None:
        rows=worksheet.nrows
    else:
        rows=question_no+1
    columns=worksheet.ncols
    question_list=[]
    for i in range(1,rows):
        q=[]
        for j in range(columns):
            q.append(worksheet.cell_value(i,j))
        l1=['a','b','c','d']
        l2=q[-1].split(',')
        d=dict(zip(l1,l2))
        q[-1]=d
        question_list.append(q)
    return question_list

easy_set=Question_set(workbook,'Easy',5)
medium_set=Question_set(workbook,'Medium')
hard_set=Question_set(workbook,'Hard')

question_set=easy_set+medium_set+hard_set

class Quesion:
    def __init__(self,question,answer,options={'a':'True','b':'False','c':'Both A and B','d':'No one'}):
        self.question=question
        self.answer=answer
        self.options=options


def take_quiz():
    if not question_set:
        print('Sorry We Do not set any question for this quiz.....')
    else:
        score=0
        wrong=[]
        for question in question_set:
            d=Quesion(question[0],question[1],question[2])
            print('-----------*---------*--------------*-----------*----------')
            print(d.question)
            for i,j in d.options.items():
                print(f'{i}: {j}')
            ans=input('Your Answer Please(a/b/c/d):- ')
            l1=['a','b','c','d']
            while ans not in l1:
                print('Kindly Choose any one option (a/b/c/d)')
                ans=input('Your Answer Please(a/b/c/d):- ')
            for i in d.options:
                if d.options[ans]==d.answer:
                    score+=1
                    break
                else:
                    wrong.append(question)
                    break
        print('-----------*---------*--------------*-----------*----------')
        print('-----------*---------*--------------*-----------*----------')
        print(f'Dear {Firstname+" "+Lastname} your score is:- {score}/{len(question_set)}')
        if not wrong:
            print('            Congratulations Dear Your All Answer Are Correct..')
            print('-----------*---------*--------------*-----------*----------')
            print('-----------*---------*--------------*-----------*----------')
        else:
            print(f'The Correct Answers For Your Questions are:- ')
            for i in wrong:
                print(f'              "{i[0]}" and Correct answer for that is "{i[1]}"')
                print('-----------*---------*--------------*-----------*----------')
            print('Please complete your fundamentals and revisit. Thank You.')

Firstname=input('Your FirstName Please:- ')
Lastname=input('Your LastName Please:- ')
want=input('Do you want to take this quiz? (y/n):- ')
if want=='y':
    take_quiz()
else:
    print('-----------*---------*--------------*-----------*----------')
    print(f'Its really an Amazing Quiz For Python Students Try it once Mr. {Firstname+" "+Lastname}')
    print('-----------*---------*--------------*-----------*----------')



