# MS-Excel-Project
I have used MS-Excel for completing 5 objectives. This Excel project simulates real-world administrative and data handling tasks. It includes payroll processing, utility billing, and grade calculation using various Excel functions and formulas to perform specific calculations. 
### Objective 1: Employee Payroll Sheet
Simulates a basic payroll system using formulas to calculate employee compensation.
- **Medical Allowance** = 7% of Basic Pay  
=Basic Pay*7%
- **House Rent Allowance** = 45% of Basic Pay  
=Basic Pay*45%
- **Gross Pay** = Basic + Allowances  
=Basic Pay + Medical Allowance + House Rent
- **Tax Calculation** (Conditional IF):  
=IF(Gross Pay > 15000, Gross Pay * 5%, Gross Pay * 3%)
- **Net Pay** = Gross Pay – Tax  
=Gross Pay - Tax
- **Grade Assignment**:  
=IF(Net Pay > 15000, "Grade-1", "Grade-2")
**Objective 2: Electricity Utility Bill**
Generates an electricity bill based on units consumed and calculates applicable charges.
- **Units Consumed** 
=IMSUB(number1, number2) where no. 1 is current reading & no. 2 is previous reading
- **Electricity Charges** Units × Rs. 2.06  
=Units*2.06
- **Surcharge (15%)**:  
=Electricity Charges*15%
- **Total Amount Due**:  
=Electricity Charges + Surcharge
- **Rounded Amount** (One decimal):  
=ROUND(Total Amount, 1)
**Objective 3: Sui Gas Utility Bill**
- **Units Consumed** = New Reading - Old Reading  
=New reading - Old reading
- **Gas Charges (Conditional Rate)**:  
=IF(Units < 200, Units*1.25, Units*1.80)
- **Sales Tax (15%)**:  
=Gas_Charges * 15%
- **Total Amount Due**:  
=Gas_Charges + Sales Tax
**Objective 4: Student Marks Certificate**
Creates a student mark sheet to calculate total marks, percentage, and assign grades.
- **Total Marks** = Sum of subject scores  
=SUM(Subject1:SubjectN)
- **Percentage**:  
=Total Marks/Maximum Marks*100
- **Grade Assignment (Nested IF)**:  
=IF(Percentage >= 70, "A", IF(Percentage >= 60, "B", "C"))

**Objective 5: Monthly Financial Summary**
- **Currency Formatting** with 2 decimal places
- **Total for Each Month** using `=SUM()`
- **Statistics for January**:
  - **Average**: =AVERAGE(A4:A7)
  - **Maximum**: =MAX(A4:A7)
  - **Minimum**: =MIN(A4:A7)
