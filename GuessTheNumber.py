from random import randint

number=randint(1,100)

score=100

while True:

    GivenNumber=int(input("Guess the number (1-100) : "))
    if GivenNumber == number:
        print("correct !")
        print("Your Score : ",score)
    elif GivenNumber> number:
        print("Think about a smaller Number")
        score-=1
    else:
        print("Think about a larger Number ")
        score-=1
