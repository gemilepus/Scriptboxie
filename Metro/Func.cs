using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Metro
{
    public class Func
    {
        public int Salary;//設定一個叫Salary的欄位

        public Func(int annualSalary)//第一個建構函式，帶入一個int參數，初始化時會將參數賦值給Salary
        {
            Salary = annualSalary;
        }

        public Func(int weeklySalary, int numberOfWeeks)//第二個建構函式，帶入2個int參數，初始化時會將參數第一個參數與第二個參數相乘後，賦值給Salary
        {
            Salary = weeklySalary * numberOfWeeks;
        }
    }
}
