using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.Wrapper
{
    public class Range
    {
        public static Range operator * (Range R, int OffSet)
        {
            OffSet += 1;
            Range rNew = new Range();
            rNew.Top = R.Top; rNew.Bottom = R.Bottom;
            rNew.Left = R.Right + OffSet;
            rNew.Right = rNew.Left + (R.Right - R.Left);

            return rNew;
        }

        public static Range operator / (Range R, int OffSet)
        {
            OffSet += 1;
            Range rNew = new Range();
            rNew.Top = R.Top; rNew.Bottom = R.Bottom;
            rNew.Right = R.Left - OffSet;
            rNew.Left = rNew.Right - (R.Right - R.Left);

            return rNew;
        }

        public static Range operator + (Range R, int OffSet)
        {
            OffSet += 1;
            Range rNew = new Range();
            rNew.Left = R.Left; rNew.Right = R.Right;
            rNew.Top = R.Bottom + OffSet;
            rNew.Bottom = rNew.Top + (R.Bottom - R.Top);

            return rNew;
        }

        public static Range operator - (Range R, int OffSet)
        {
            OffSet += 1;
            Range rNew = new Range();
            rNew.Left = R.Left; rNew.Right = R.Right;
            rNew.Bottom = R.Top - OffSet;
            rNew.Top = rNew.Bottom - (R.Bottom - R.Top);

            return rNew;
        }

        public int Top { get; internal set; }
        public int Bottom { get; internal set; }
        public int Left { get; internal set; }
        public int Right { get; internal set; }

        public static bool IsRowShift(Range R1, Range R2)
        {
            if (R1.Bottom > R2.Top)
            {
                return true;
            }
            else if (R2.Bottom > R1.Top)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public static bool IsColShift(Range R1, Range R2)
        {
            if (R1.Right < R2.Left)
            {
                return true;
            }
            else if (R2.Right < R1.Left)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public static bool IsShift(Range R1, Range R2)
        {
            return Range.IsRowShift(R1, R2) && Range.IsColShift(R1, R2);
        }

        public override string ToString()
        {
            return 
                string.Format(
                "{T:{0},B:{1},L:{2},R:{3}}",
                this.Top, this.Bottom, this.Left, this.Right
                );
        }
    }
}
