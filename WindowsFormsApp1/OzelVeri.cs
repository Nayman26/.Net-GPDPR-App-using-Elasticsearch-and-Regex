using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApp1
{
    class OzelVeri
    {
        private string token;
        private int start_offset;
        private int end_offset;
        private int position;

        public int Position { get => position; set => position = value; }
        public string Token { get => token; set => token = value; }
        public int Start_offset { get => start_offset; set => start_offset = value; }
        public int End_offset { get => end_offset; set => end_offset = value; }
    }
}
