using System.Collections.Generic;
using System.Linq;

namespace DayBillSummary.Extensions
{
    public static class Extensions
    {
        public static string OrderNumberCombine(this IEnumerable<int> source)
        {
            var orderNumbers = source.ToArray();

            var max = orderNumbers.Max();
            var min = orderNumbers.Min();
            var result = "";
            int index = 0;

            for (int i = min; i < max; i++)
            {
                if (orderNumbers[index] != i)
                {
                    result = orderNumbers[index - 1] == min ? $"{result}{min}," : $"{result}{min}-{orderNumbers[index - 1]},";
                    min = orderNumbers[index];
                    i = min;
                }
                index++;
            }

            result = min == max ? $"{result}{min}" : $"{result}{min}-{max}";
            return result;
        }
    }
}