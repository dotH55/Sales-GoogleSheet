
// This class is used to facilitate updating the Google-Spreadsheets
// Variables: name, currYearMonth[] & priorYearMonth[]
// The two arrays store monthly sales for the current year and prior year
// It also has getters & setters for them

using System;
using System.Collections.Generic;

namespace SalesGoogleSheet
{
    class Employee
    {
        string name;
        // Double dimension arrays for
        // Monthly sales: i & monthly units sold: j
        private double[,] currentYear = new double[13, 2];
        private double[,] priorYear = new double[13, 2];

        public Employee(string name)
        {
            // Initializes arrays to zero
            this.name = name;
            for (int i = 0; i < 13; i++){
                for (int j = 0; j < 2; j++){

                    currentYear[i, j] = 0;
                    priorYear[i, j] = 0;
                }
            }
        } // Employee

        // Getters & Setters
        public string GetName()
        {
            return name;
        }
        public void SetName(string name)
        {
            this.name = name;
        }

        public double GetCurrYearMonthlySales(int month)
        {
            return currentYear[month, 0];
        }

        public void SetCurrYearMonthlySales(int month, double sales)
        {
            currentYear[month, 0] = sales;
        }

        public int GetCurrYearMonthlyUnitsSold(int month)
        {
            return (int)currentYear[month, 1];
        }

        public void SetCurrYearMonthlyUnitsSold(int month, int units)
        {
            currentYear[month, 1] = units;
        }

        public double GetPriorYearMonthlySales(int month)
        {
            return priorYear[month, 0];
        }

        public void SetPriorYearMonthlySales(int month, double sales)
        {
            priorYear[month, 0] = sales;
        }

        public int GetPriorYearMonthlyUnitsSold(int month)
        {
            return (int)priorYear[month, 1];
        }

        public void SetPriorYearMonthlyUnitsSold(int month, int units)
        {
            priorYear[month, 1] = units;
        }

        public double GetPriorYearQuarterlySales(int qtr)
        {
            double quarter = 0;
            qtr = (qtr * 3) - 2;
            for (int j = qtr; j <= (qtr + 2); j++)
            {
                quarter += GetPriorYearMonthlySales(j);
            }
            return quarter;
        } // getPriorQuarter

        public double GetPriorYearQuarterlyUnitsSold(int qtr)
        {
            double quarter = 0;
            qtr = (qtr * 3) - 2;
            for (int j = qtr; j <= (qtr + 2); j++)
            {
                quarter += GetPriorYearMonthlyUnitsSold(j);
            }
            return quarter;
        } // getPriorQuarter

        public double GetCurrYearQuaterlySales(int qtr)
        {
            double quarter = 0;
            qtr = (qtr * 3) - 2;
            for (int j = qtr; j <= (qtr + 2); j++)
            {
                quarter += GetCurrYearMonthlySales(j);
            }
            return quarter;
        } // getCurrYearQtr

        public double GetCurrYearQuaterlyUnitsSold(int qtr)
        {
            double quarter = 0;
            qtr = (qtr * 3) - 2;
            for (int j = qtr; j <= (qtr + 2); j++)
            {
                quarter += GetCurrYearMonthlyUnitsSold(j);
            }
            return quarter;
        } // getCurrYearQtr

        public double GetPriorYearQuaterlySales(List<Employee> employee, int qtr)
        {
            double quarter = 0;
            foreach(Employee i in employee)
            {
                quarter += i.GetPriorYearQuarterlySales(qtr);
            }
            return quarter;
        } // getPriorYearQtr

        public double GetPriorYearQuaterlyUnitsSold(List<Employee> employee, int qtr)
        {
            double quarter = 0;
            foreach (Employee i in employee)
            {
                quarter += i.GetPriorYearQuarterlyUnitsSold(qtr);
            }
            return quarter;
        } // getPriorYearQtr

        public double GetCurrYearQuaterlySales(List<Employee> employee, int qtr)
        {
            double quarter = 0;
            foreach (Employee i in employee)
            {
                quarter += i.GetCurrYearQuaterlySales(qtr);
            }
            return quarter;
        } // getCurrYearQtr

        public double GetCurrYearQuaterlyUnitsSold(List<Employee> employee, int qtr)
        {
            double quarter = 0;
            foreach (Employee i in employee)
            {
                quarter += i.GetCurrYearQuaterlyUnitsSold(qtr);
            }
            return quarter;
        } // getCurrYearQtr

        public double GetPriorYearTotalSales(List<Employee> employee)
        {
            double total = 0;
            for(int i = 1; i <= 4; i++)
            {
                total += GetPriorYearQuaterlySales(employee, i);
            }
            return total;
        } // getPriorYearTotal

        public double GetPriorYearTotalUnitsSold(List<Employee> employee)
        {
            double total = 0;
            for (int i = 1; i <= 4; i++)
            {
                total += GetPriorYearQuaterlyUnitsSold(employee, i);
            }
            return total;
        } // getPriorYearTotal

        public double GetCurrYearTotalSales(List<Employee> employee)
        {
            double total = 0;
            for (int i = 1; i <= 4; i++)
            {
                total += GetCurrYearQuaterlySales(employee, i);
            }
            return total;
        } // getCurrYearTotal

        public double GetCurrYearTotalUnitsSold(List<Employee> employee)
        {
            double total = 0;
            for (int i = 1; i <= 4; i++)
            {
                total += GetCurrYearQuaterlyUnitsSold(employee, i);
            }
            return total;
        } // getCurrYearTotal

        // This print statement is for debugging purposes
        public void PrintEmployee()
        {
            Console.WriteLine("Name: " + GetName());
            for (int i = 1; i <= 12; i++)
            {
                Console.WriteLine("Prior Year Units Sold: " + GetPriorYearMonthlyUnitsSold(i) + " Current Year Sales: " + GetCurrYearMonthlyUnitsSold(i));
                Console.WriteLine("Prior Year Sales: " + GetPriorYearMonthlySales(i) + " Current Year Sales: " + GetCurrYearMonthlySales(i));
            } // for
        } // printEmployee
    } // Employee
} // TestGoogleSS
