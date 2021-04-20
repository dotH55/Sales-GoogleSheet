using System;

// This class is used to facilitate reading data obtain from the DSI server
// Variables: year, month, name & monthly sales
// Also has getters & setters for each variable

namespace SalesGoogleSheet
{
    class Entry
    {
        private int year;
        private int month;
        private string name;
        private double monthlySales;
        private int unitsSold;

        public Entry(int year, int month, string name, double monthlySales, int unitsSold)
        {
            this.year = year;
            this.month = month;
            this.name = name;
            this.monthlySales = monthlySales;
            this.unitsSold = unitsSold;
        }

        // Getters & Setters
        public int GetYear()
        {
            return year;
        }

        public void SetYear(int year)
        {
            this.year = year;
        }

        public int GetMonth()
        {
            return month;
        }

        public void SetMonth(int month)
        {
            this.month = month;
        }

        public string GetName()
        {
            return name;
        }

        public void SetName(String name)
        {
            this.name = name;
        }

        public double GetMonthlySales()
        {
            return monthlySales;
        }

        public void SetMonthlySales(double monthlySales)
        {
            this.monthlySales = monthlySales;
        }
        public int GetMonthlyUnitsSold()
        {
            return unitsSold;
        }

        public void SetMonthlyUnitsSold(int unitsSold)
        {
            this.unitsSold = unitsSold;
        }

        // This print method is for debugging purposes
        public void PrintEntry()
        {
            Console.WriteLine("Year: " + GetYear() + " Month: " + GetMonth() + " Employee: " + GetName() + " Sales: " + GetMonthlySales() + "Units Sold: " + GetMonthlyUnitsSold());
        } // printEntry
    } // Entry
} // SalesGoogleSheet
