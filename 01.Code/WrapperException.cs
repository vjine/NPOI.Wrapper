using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.Wrapper
{
    public class WrapperException : Exception
    {
        public WrapperException(string Msg, params object[] args)
            : base(string.Format(Msg, args))
        {

        }
    }
}
