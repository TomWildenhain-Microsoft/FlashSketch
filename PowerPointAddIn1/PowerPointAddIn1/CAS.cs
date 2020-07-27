using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.NetworkInformation;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointAddIn1
{
    class CAS
    {
        public static CasVar NewVar()
        {
            throw new Exception("Not implemented");
        }
        public static CasExpr VarExpr()
        {
            throw new Exception("Not implemented");
        }
        public static CasExpr ConstExpr(int c)
        {
            throw new Exception("Not implemented");
        }
        public static CasExpr Add(CasExpr expr1, CasExpr expr2)
        {
            throw new Exception("Not implemented");
        }
        public static CasExpr Sub(CasExpr expr1, CasExpr expr2)
        {
            throw new Exception("Not implemented");
        }
        public static CasExpr Mul(CasExpr expr1, CasExpr expr2)
        {
            throw new Exception("Not implemented");
        }
        public static CasExpr Div(CasExpr expr1, CasExpr expr2)
        {
            throw new Exception("Not implemented");
        }

        // Result is null if solver fails
        public static Tuple<CasExpr, CasVar> SolveFor(CasExpr expr, CasVar v)
        {
            throw new Exception("Not implemented");
        }
        public static float Eval(CasExpr expr, Dictionary<CasVar, float> values)
        {
            throw new Exception("Not implemented");
        }
    }

    class CasExpr
    {

    }

    class CasVar
    {

    }
}
