using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ParseHTML
{
    class progress_bar
    {
        public int min { get; set; }
        public int max { get; set; }
        private int current { get; set; }

        public progress_bar(int my_min, int my_max)
        {
            min = my_min;
            max = my_max;
        }
        public progress_bar() { }

        private int calc_percent ()
        {
            if (current==min)
            {
                return 0;
            }
            else if (current==max)
            {
                return 100;
            }
            else
            {
                int result = (int)Math.Floor((double)((current * 100) / (max - min)));
                return result;
            }
        }

        public void print_progressBar (int my_cur_pos)
        {
            current = my_cur_pos;
            
            int current_percent = calc_percent();
            int current_symbols = (int)Math.Floor((double)(current_percent / 2));

            Console.Write("[");

            for (int i = 0; i < current_symbols; i++)
            {
                Console.Write("#");
            }

            for (int i = current_symbols; i < 50; i++)
            {
                Console.Write(".");
            }
            Console.Write(String.Format("]\t{0}%",current_percent));
        }
    }
}
