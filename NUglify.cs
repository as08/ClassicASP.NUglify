using System;
using System.Runtime.InteropServices;
using NUglify;

namespace NF
{

    [ProgId("ClassicASP.NUglify")]

    [ClassInterface(ClassInterfaceType.AutoDual)]

    [Guid("027F1577-9D93-4699-B323-261296F822F1")]

    [ComVisible(true)]
    public class NUglify
    {

        [ComVisible(true)]

        public string JS(string jsStr)
        {
            var result = Uglify.Js(jsStr);
            return result.Code;
        }

        public string CSS(string cssStr)
        {
            var result = Uglify.Css(cssStr);
            return result.Code;
        }

    }

}