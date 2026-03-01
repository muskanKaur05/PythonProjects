# Input we needed from user
# Total rent paid by user
# Total food ordered for snacking
# Electricity units spend
# Charge per unit

# Output
# Total amount you have paid

rent = int(input("Please enter your total rent : "))
food = int(input("Please enter your total food expense : "))
unitspend = int(input("Enter the electricity unit spent : "))
chargePunit = int(input("Enter the charge per unit : "))
persons = int(input("Enter the number of persons : "))
totalamount = rent + food + (unitspend*chargePunit)
perpersonamount= totalamount/persons
print("Per person cost : ","Rs. ",perpersonamount)
