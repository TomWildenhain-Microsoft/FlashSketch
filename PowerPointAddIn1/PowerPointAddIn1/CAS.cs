using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Net.NetworkInformation;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointAddIn1
{
    class CasSystem
    {
        public static CasSystem Instance = new CasSystem();
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
                return new CasExpr(CasPolynomial.PolySum(expr1.Poly1, expr2.Poly1), expr2.Poly2).Simplify();
            }
            else
            {
                var p1 = CasPolynomial.PolyProd(expr1.Poly1, expr2.Poly2);
                var p2 = CasPolynomial.PolyProd(expr1.Poly2, expr2.Poly1);
                var num = CasPolynomial.PolySum(p1, p2);
                var denom = CasPolynomial.PolyProd(expr1.Poly2, expr2.Poly2);
                return new CasExpr(num, denom).Simplify();
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
                return new CasExpr(expr2.Poly1, expr1.Poly2).Simplify();
            }
            if (expr1.Poly2.PolyEquals(expr2.Poly1))
            {
                return new CasExpr(expr1.Poly1, expr2.Poly2).Simplify();
            }
            return new CasExpr(CasPolynomial.PolyProd(expr1.Poly1, expr2.Poly1), CasPolynomial.PolyProd(expr1.Poly2, expr2.Poly2)).Simplify();
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
                if (!t.Variables.ContainsKey(v))
                {
                    zeroTerms.Add(t);
                }
                else if (t.Variables[v] == 1)
                {
                    oneTerms.Add(t);
                }
                else
                {
                    return null;
                }
            }
            if (oneTerms.Count == 0) return null;
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
                t.Variables.Remove(v);
            }
            return new CasExpr(poly1, poly2);
        }

        // Result is null if solver fails
        public Tuple<CasExpr, CasVar> Solve(CasExpr expr, CasVar forV)
        {
            if (expr.Poly1.ContainsVar(forV) && expr.Poly2.ContainsVar(forV))
            {
                return null;
            }
            if (!expr.Poly1.ContainsVar(forV) && !expr.Poly2.ContainsVar(forV))
            {
                return null;
            }
            if (expr.Poly1.ContainsVar(forV) && !expr.Poly2.ContainsVar(forV))
            {
                List<CasTerm> zeroTerms = new List<CasTerm>();
                List<CasTerm> oneTerms = new List<CasTerm>();
                foreach (var t in expr.Poly1.Terms)
                {
                    if (!t.Variables.ContainsKey(forV))
                    {
                        zeroTerms.Add(t);
                    }
                    else if (t.Variables[forV] == 1)
                    {
                        oneTerms.Add(t);
                    }
                    else
                    {
                        return null;
                    }
                }
                if (oneTerms.Count == 0) return null;
                CasPolynomial polyZeros = new CasPolynomial();
                CasPolynomial polyOnes = new CasPolynomial();
                foreach (var t in zeroTerms)
                {
                    polyZeros.Terms.Add(t.CopyWithCoefficient(-1 * t.Coefficient));
                }
                foreach (var t in oneTerms)
                {
                    polyOnes.Terms.Add(t.CopyWithCoefficient(t.Coefficient));
                    t.Variables.Remove(forV);
                }
                var newVar = NewVar();
                var newTerm = new CasTerm(1);
                newTerm.Variables[newVar] = 1;
                polyZeros.Terms.Add(newTerm);
                var resExpr = new CasExpr(CasPolynomial.PolyProd(polyZeros, expr.Poly2), polyOnes);
                return new Tuple<CasExpr, CasVar>(resExpr, newVar);
            }
            else
            {
                List<CasTerm> zeroTerms = new List<CasTerm>();
                List<CasTerm> oneTerms = new List<CasTerm>();
                foreach (var t in expr.Poly2.Terms)
                {
                    if (!t.Variables.ContainsKey(forV))
                    {
                        zeroTerms.Add(t);
                    }
                    else if (t.Variables[forV] == 1)
                    {
                        oneTerms.Add(t);
                    }
                    else
                    {
                        return null;
                    }
                }
                if (oneTerms.Count == 0) return null;
                CasPolynomial polyZeros = new CasPolynomial();
                CasPolynomial polyOnes = new CasPolynomial();
                var newVar = NewVar();
                foreach (var t in zeroTerms)
                {
                    polyZeros.Terms.Add(t.CopyWithCoefficient(-1 * t.Coefficient));
                    t.Variables[newVar] = 1;
                }
                foreach (var t in oneTerms)
                {
                    polyOnes.Terms.Add(t.CopyWithCoefficient(t.Coefficient));
                    t.Variables.Remove(forV);
                    t.Variables[newVar] = 1;
                }
                var resExpr = new CasExpr(CasPolynomial.PolySum(polyZeros, expr.Poly2), polyOnes);
                return new Tuple<CasExpr, CasVar>(resExpr, newVar);
            }
        }

        public CasExpr Substitute(CasExpr expr, CasVar forV, CasExpr inExpr)
        {
            CasExpr expr1 = inExpr.Poly1.Substitute(expr, forV);
            CasExpr expr2 = inExpr.Poly2.Substitute(expr, forV);
            return Div(expr1, expr2);
        }
        public float Eval(CasExpr expr, Dictionary<CasVar, float> values)
        {
            float res1 = expr.Poly1.Eval(values);
            float res2 = expr.Poly2.Eval(values);
            return res1 / res2;
        }

        public HashSet<CasVar> UsedVariables(CasExpr expr)
        {
            HashSet<CasVar> res = new HashSet<CasVar>();
            expr.Poly1.GetUsedVariables(res);
            expr.Poly2.GetUsedVariables(res);
            return res;
        }

        public static int GCD(int a, int b)
        {
            while (a != 0 && b != 0)
            {
                if (Math.Abs(a) > Math.Abs(b))
                    a %= b;
                else
                    b %= a;
            }
            return a | b;
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
            var copy = new CasTerm(c);
            foreach (var v in Variables)
            {
                copy.Variables[v.Key] = v.Value;
            }
            return copy;
        }

        public float Eval(Dictionary<CasVar,float> values) 
        {
            float res = 1.0f;
            foreach (var v in Variables.Keys) {
                res *= (float)Math.Pow(values[v], Variables[v]);
            }
            return res * Coefficient; 
        }

        public CasExpr Substitute(CasExpr expr, CasVar forV)
        {
            if (!Variables.ContainsKey(forV))
            {
                return this.ToExpr();
            }
            var exprPow = expr.RaisedToPower(Variables[forV]);
            var newTerm = CopyWithCoefficient(Coefficient);
            newTerm.Variables.Remove(forV);
            return CasSystem.Instance.Mul(newTerm.ToExpr(), exprPow);
        }

        public CasExpr ToExpr()
        {
            var p = new CasPolynomial();
            p.Terms.Add(this);
            return new CasExpr(p, CasPolynomial.ConstantPoly(1));
        }

        public void Simplify()
        {
            foreach (var v in new List<CasVar>(Variables.Keys))
            {
                if (Variables[v] == 0)
                {
                    Variables.Remove(v);
                }
            }
        }
        public void GetUsedVariables(HashSet<CasVar> res)
        {
            foreach (var v in Variables.Keys)
            {
                res.Add(v);
            }
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
            res.Simplify();
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
            if (c != 0)
            {
                p.Terms.Add(new CasTerm(c));
            }
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

        public float Eval(Dictionary<CasVar, float> values)
        {
            float res = 0.0f;
            foreach (var t in Terms)
            {
                res += t.Eval(values);
            }
            return res;
        }

        public int TermGCD()
        {
            int res = 0;
            foreach (var term in Terms)
            {
                res = CasSystem.GCD(res, term.Coefficient);
            }
            return res;
        }

        public CasPolynomial DivideTerms(int d)
        {
            CasPolynomial res = new CasPolynomial();
            foreach (var term in Terms)
            {
                if (term.Coefficient % d != 0)
                {
                    throw new Exception("Attempt to divide polynomial by int it isn't divisible by");
                }
                res.Terms.Add(term.CopyWithCoefficient(term.Coefficient / d));
            }
            return res;
        }

        public CasExpr Substitute(CasExpr expr, CasVar forV)
        {
            if (!ContainsVar(forV)) return new CasExpr(this, CasPolynomial.ConstantPoly(1));
            List<CasExpr> exprs = new List<CasExpr>();
            foreach (var t in Terms)
            {
                exprs.Add(t.Substitute(expr, forV));
            }
            var sum = CasExpr.ConstantExpr(0);
            foreach (var e in exprs)
            {
                sum = CasSystem.Instance.Add(sum, e);
            }
            return sum;
        }

        public bool ContainsVar(CasVar v)
        {
            foreach (var t in Terms)
            {
                if (t.Variables.ContainsKey(v)) return true;
            }
            return false;
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
        public void GetUsedVariables(HashSet<CasVar> res)
        {
            foreach (var t in Terms)
            {
                t.GetUsedVariables(res);
            }
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
                sep = "+";
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
            if (poly2.Terms.Count == 0)
            {
                throw new Exception("Attempt to divide by zero.");
            }
        }
        public CasExpr RaisedToPower(int exp)
        {
            if (exp == 0) return ConstantExpr(1);
            CasExpr res = this;
            for (int i = 1; i < exp; i++)
            {
                res = CasSystem.Instance.Mul(res, this);
            }
            return res;
        }
        public static CasExpr ConstantExpr(int c)
        {
            return new CasExpr(CasPolynomial.ConstantPoly(c), CasPolynomial.ConstantPoly(1));
        }
        public override string ToString()
        {
            return Poly1.ToString() + "//" + Poly2.ToString();
        }
        public CasExpr Simplify()
        {
            int gcd = CasSystem.GCD(Poly1.TermGCD(), Poly2.TermGCD());
            if (gcd == 0 || gcd == 1)
            {
                return this;
            }
            else
            {
                return new CasExpr(Poly1.DivideTerms(gcd), Poly2.DivideTerms(gcd));
            }
        }
    }

    class CasVar : IComparable<CasVar>
    {
        public string Name;
        public CasVar(string name)
        {
            Name = name;
        }
        public override string ToString()
        {
            return Name;
        }
        public int CompareTo(CasVar other)
        {
            return Name.CompareTo(other.Name);
        }
    }
}
