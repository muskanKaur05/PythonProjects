"""
FLOW:
1. User will input from (Rocke, Paper, Scissor)
2. Computer will pick random value from (Rock, Paper, Scissor)
3. Result Print

Cases:

A - Rock
Rock - Rock = Tie
Rock - Paper = Paper Win
Rock - Scissor = Rock Win

B - Paper
Paper - Paper = Tie
Paper - Rock = Paper Win
Paper - Scissor = Scissor Win

C - Scissor
Scissor - Scissor = Tie
Scissor - Paper = Scissor Win
Scissor - Rock = Rock Win
"""

import random
User_choice= input ("Enter your choice from Rock/ Paper/ Scissor : ")
input_choices=["Rock","Scissor","Paper"]
computer_choice=random.choice(input_choices)

print(f'User Choice = {User_choice}, Computer Choice = {computer_choice}')


if User_choice==computer_choice:
    print("Both chooses same : Match Tie")
else:
    if User_choice=='Rock':
        if computer_choice=='Paper':
            print("Paper covers Rock, Computer Win")
        else:
            print("Rock cuts Scissor, You Win")
    elif User_choice=="Paper":
        if computer_choice=="Rock":
            print("Paper covers Rock, You Win")
        else:
            print("Scissor cuts Paper, Computer Win")
    else:
        if computer_choice=="Paper":
            print("Scissor cuts Paper, You Win")
        else:
            print("Rock cuts Scissor, Computer Win")
            

    
    

