using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Net.NetworkInformation;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointAddIn1
{
    class CasSystem
    {
        private int NumVars = 0;
        public CasVar NewVar()
        {
            var v = new CasVar("Var" + NumVars);
            NumVars++;
            return v;
        }
        public CasExpr VarExpr(CasVar v)
        {
            var t = new CasTerm(1);
            t.Variables[v] = 1;
            var p = new CasPolynomial();
            p.Terms.Add(t);
            return new CasExpr(p, CasPolynomial.ConstantPoly(1));
        }
        public CasExpr ConstExpr(int c)
        {
            return CasExpr.ConstantExpr(c);
        }
        public CasExpr Add(CasExpr expr1, CasExpr expr2)
        {
            if (expr1.Poly2.PolyEquals(expr2.Poly2))
            {
                return new CasExpr(CasPolynomial.PolySum(expr1.Poly1, expr2.Poly1), expr2.Poly2);
            }
            else
            {
                var p1 = CasPolynomial.PolyProd(expr1.Poly1, expr2.Poly2);
                var p2 = CasPolynomial.PolyProd(expr1.Poly2, expr2.Poly1);
                var num = CasPolynomial.PolySum(p1, p2);
                var denom = CasPolynomial.PolyProd(expr1.Poly2, expr2.Poly2);
                return new CasExpr(num, denom);
            }
        }
        public CasExpr Sub(CasExpr expr1, CasExpr expr2)
        {
            return Add(expr1, Mul(expr2, ConstExpr(-1)));
        }
        public CasExpr Mul(CasExpr expr1, CasExpr expr2)
        {
            if (expr1.Poly1.PolyEquals(expr2.Poly2))
            {
                return new CasExpr(expr2.Poly1, expr1.Poly2);
            }
            if (expr1.Poly2.PolyEquals(expr2.Poly1))
            {
                return new CasExpr(expr1.Poly1, expr2.Poly2);
            }
            return new CasExpr(CasPolynomial.PolyProd(expr1.Poly1, expr2.Poly1), CasPolynomial.PolyProd(expr1.Poly2, expr2.Poly1));
        }
        public CasExpr Div(CasExpr expr1, CasExpr expr2)
        {
            return Mul(expr1, new CasExpr(expr2.Poly2, expr2.Poly1));
        }

        // Result is null if solver fails
        public CasExpr MakeZero(CasExpr expr, CasVar v)
        {
            List<CasTerm> zeroTerms = new List<CasTerm>();
            List<CasTerm> oneTerms = new List<CasTerm>();
            foreach (var t in expr.Poly1.Terms)
            {
                if (t.Variables[v] == 0)
                {
                    zeroTerms.Add(t);
                }
                else if (t.Variables[v] == 1)
                {
                    oneTerms.Add(t);
                }
            }
            CasPolynomial poly1 = new CasPolynomial();
            CasPolynomial poly2 = new CasPolynomial();
            foreach (var t in zeroTerms)
            {
                poly1.Terms.Add(t.CopyWithCoefficient(-1 * t.Coefficient));
            }
            foreach (var t in oneTerms)
            {
                poly2.Terms.Add(t.CopyWithCoefficient(t.Coefficient));
            }
            foreach (var t in poly2.Terms)
            {
                t.Variables[v]--;
            }
            return new CasExpr(poly1, poly2);
        }

        public CasExpr Substitute(CasExpr expr, CasVar forV, CasExpr inExpr)
        {
            throw new Exception("Not implemented");
        }
        public float Eval(CasExpr expr, Dictionary<CasVar, float> values)
        {
            throw new Exception("Not implemented");
        }
    }

    class CasTerm
    {
        public int Coefficient;
        public Dictionary<CasVar, int> Variables;
        public CasTerm(int coefficient)
        {
            Coefficient = coefficient;
            Variables = new Dictionary<CasVar, int>();
        }
        public string StringSuffix()
        {
            List<CasVar> vars = new List<CasVar>(Variables.Keys);
            vars.Sort();
            var res = "";
            foreach (var v in vars)
            {
                res += "*" + v.Name;
            }
            return res;
        }
        override public string ToString()
        {
            return Coefficient + StringSuffix();
        }
        public CasTerm CopyWithCoefficient(int c)
        {
            if (c == Coefficient) return this;
            var copy = new CasTerm(c);
            foreach (var v in Variables)
            {
                copy.Variables[v.Key] = v.Value;
            }
            return copy;
        }
        public static CasTerm TermProduct(CasTerm term1, CasTerm term2)
        {
            var res = new CasTerm(term1.Coefficient * term2.Coefficient);
            foreach (var v in term1.Variables.Keys)
            {
                if (res.Variables.ContainsKey(v))
                {
                    res.Variables[v] += term1.Variables[v];
                }
                else
                {
                    res.Variables[v] = term1.Variables[v];
                }
            }
            foreach (var v in term2.Variables.Keys)
            {
                if (res.Variables.ContainsKey(v))
                {
                    res.Variables[v] += term2.Variables[v];
                }
                else
                {
                    res.Variables[v] = term2.Variables[v];
                }
            }
            return res;
        }
    }
    
    class CasPolynomial
    {
        public List<CasTerm> Terms;
        public CasPolynomial()
        {
            Terms = new List<CasTerm>();
        }
        public static CasPolynomial ConstantPoly(int c)
        {
            var p = new CasPolynomial();
            p.Terms.Add(new CasTerm(c));
            return p;
        }
        public bool PolyEquals(CasPolynomial poly)
        {
            HashSet<string> terms = new HashSet<string>();
            foreach (var t in poly.Terms)
            {
                terms.Add(t.ToString());
            }
            foreach (var t in Terms)
            {
                if (!terms.Contains(t.ToString()))
                {
                    return false;
                }
                terms.Remove(t.ToString());
            }
            return terms.Count == 0;
        }
        public void Simplify()
        {
            Dictionary<string, Tuple<CasTerm, int>> termsDict = new Dictionary<string, Tuple<CasTerm, int>>();
            var newTerms = new List<CasTerm>();
            foreach (var t in Terms)
            {
                string s = t.StringSuffix();
                int c = t.Coefficient;
                if (termsDict.ContainsKey(s))
                {
                    c += termsDict[s].Item2;
                }
                termsDict[t.StringSuffix()] = new Tuple<CasTerm, int>(t, c);
            }
            foreach (var t in termsDict.Values)
            {
                var newTerm = t.Item1.CopyWithCoefficient(t.Item2);
                if (newTerm.Coefficient != 0)
                {
                    newTerms.Add(newTerm);
                }
            }
            Terms = newTerms;
        }
        public static CasPolynomial PolySum(CasPolynomial poly1, CasPolynomial poly2)
        {
            CasPolynomial res = new CasPolynomial();
            foreach (var t in poly1.Terms)
            {
                res.Terms.Add(t);
            }
            foreach (var t in poly2.Terms)
            {
                res.Terms.Add(t);
            }
            res.Simplify();
            return res;
        }
        public static CasPolynomial PolyProd(CasPolynomial poly1, CasPolynomial poly2)
        {
            CasPolynomial res = new CasPolynomial();
            foreach (var t1 in poly1.Terms)
            {
                foreach (var t2 in poly2.Terms)
                {
                    res.Terms.Add(CasTerm.TermProduct(t1, t2));
                }
            }
            res.Simplify();
            return res;
        }
        public override string ToString()
        {
            var res = "";
            var sep = "";
            foreach (var t in Terms)
            {
                res += sep + t.ToString();
                sep = " + ";
            }
            return res;
        }
    }
    class CasExpr
    {
        public CasPolynomial Poly1;
        public CasPolynomial Poly2;
        public CasExpr(CasPolynomial poly1, CasPolynomial poly2)
        {
            Poly1 = poly1;
            Poly2 = poly2;
        }
        public static CasExpr ConstantExpr(int c)
        {
            return new CasExpr(CasPolynomial.ConstantPoly(c), CasPolynomial.ConstantPoly(1));
        }
        public override string ToString()
        {
            return Poly1.ToString() + "   //   " + Poly2.ToString();
        }
    }

    class CasVar : IComparable<CasVar>
    {
        public string Name;
        public CasVar(string name)
        {
            Name = name;
        }
        public int CompareTo(CasVar other)
        {
            return Name.CompareTo(other.Name);
        }
    }
}
