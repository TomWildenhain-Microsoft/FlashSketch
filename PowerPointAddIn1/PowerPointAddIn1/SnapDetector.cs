using Microsoft.VisualStudio.Tools.Applications.Runtime;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointAddIn1
{
    class SnapDetector
    {
        public static SnapDetector Instance = new SnapDetector();

        Dictionary<CasVar, VariableCacheEntry> Variables = new Dictionary<CasVar, VariableCacheEntry>();
        HashSet<SnapCacheEntry> CacheEntries = new HashSet<SnapCacheEntry>();
        Dictionary<string, SnapCacheEntry> CachedXs = new Dictionary<string, SnapCacheEntry>();
        Dictionary<string, SnapCacheEntry> CachedYs = new Dictionary<string, SnapCacheEntry>();
        Dictionary<string, SnapCacheEntry> CachedLengths = new Dictionary<string, SnapCacheEntry>();

        CasSystem CAS;
        public SnapDetector()
        {
            CAS = CasSystem.Instance;
        }

        public void EqualizeLongerDims(List<SlideObject> slideObjects, int shapeId1, int shapeId2)
        {
            UpdateSnapCache(slideObjects);
            FindExprsForSlideObjs(slideObjects);
            CasExpr expr1 = null;
            CasExpr expr2 = null;
            foreach (var obj in slideObjects)
            {
                if (obj.ShapeId == shapeId1)
                {
                    expr1 = obj.LongerDimExpr();
                }
                if (obj.ShapeId == shapeId2)
                {
                    expr2 = obj.LongerDimExpr();
                }
            }
            EqualizeExpressions(expr1, expr2, expr1, slideObjects);
        }

        public void MakeSquare(List<SlideObject> slideObjects, int shapeId)
        {
            UpdateSnapCache(slideObjects);
            FindExprsForSlideObjs(slideObjects);
            SlideObject shape = null;
            foreach (var obj in slideObjects)
            {
                if (obj.ShapeId == shapeId)
                {
                    shape = obj;
                }
            }
            EqualizeExpressions(shape.HeightExpr, shape.WidthExpr, shape.ShorterDimExpr(), slideObjects);
        }

        public void EqualizeExpressions(CasExpr expr1, CasExpr expr2, CasExpr invariant, List<SlideObject> slideObjects)
        {
            var diffExpr = CAS.Sub(expr1, expr2);
            var vars = new List<CasVar>(CAS.UsedVariables(diffExpr));
            CasVar bestVar = null;
            CasExpr bestSolution = null;
            float minChange = 0;
            Dictionary<CasVar, float> values = GetVariableValues();
            float initialVal = CAS.Eval(invariant, values);
            foreach (var v in vars)
            {
                var solution = CAS.MakeZero(diffExpr, v);
                if (solution != null)
                {
                    float change = Math.Abs(CAS.Eval(CAS.Substitute(solution, v, invariant), values) - initialVal);
                    if (bestVar == null || change < minChange)
                    {
                        bestVar = v;
                        minChange = change;
                        bestSolution = solution;
                    }
                }
            }
            if (bestSolution != null)
            {
                ReplaceVariable(bestVar, bestSolution);
            }

            FindExprsForSlideObjs(slideObjects);

            SlideScanner.Instance.ApplyDimsToShapes(slideObjects, values);
            CacheEntries = new HashSet<SnapCacheEntry>();
            CachedXs = RebuildCache(CachedXs, values);
            CachedYs = RebuildCache(CachedYs, values);
            CachedLengths = RebuildCache(CachedLengths, values);
        }

        public Dictionary<CasVar, float> GetVariableValues()
        {
            Dictionary<CasVar, float> values = new Dictionary<CasVar, float>();
            foreach (var v in Variables.Values)
            {
                values[v.Variable] = v.Value;
            }
            return values;
        }

        public Dictionary<string, SnapCacheEntry> RebuildCache(Dictionary<string, SnapCacheEntry> cache, Dictionary<CasVar, float> values)
        {
            Dictionary<string, SnapCacheEntry> res = new Dictionary<string, SnapCacheEntry>();
            foreach (var entry in cache.Values)
            {
                entry.Key = FloatToKey(CAS.Eval(entry.Expr, values));
                if (!res.ContainsKey(entry.Key))
                {
                    res[entry.Key] = entry;
                    entry.Cache = res;
                    CacheEntries.Add(entry);
                }
            }
            return res;
        }

        public void FindExprsForSlideObjs(List<SlideObject> slideObjects)
        {
            foreach (var obj in slideObjects)
            {
                obj.XExpr = CachedXs[FloatToKey(obj.X1)].Expr;
                obj.WidthExpr = CachedLengths[FloatToKey(obj.Width)].Expr;
                obj.YExpr = CachedYs[FloatToKey(obj.Y1)].Expr;
                obj.HeightExpr = CachedLengths[FloatToKey(obj.Height)].Expr;
            }
        }

        public void UpdateSnapCache(List<SlideObject> slideObjects)
        {
            FlashSketch.Instance.ClearNotes();
            foreach (var obj in slideObjects)
            {
                var resX = UpdateForDim(obj.X1, obj.CX, obj.X2, obj.Width, CachedXs);
                var resY = UpdateForDim(obj.Y1, obj.CY, obj.Y2, obj.Height, CachedYs);
            }
            RemoveUnusedVariables(slideObjects);
            foreach (var v in Variables.Values)
            {
                FlashSketch.Instance.PrintToNotes(v.Variable.Name + ": " + v.Value);
            }
        }

        public Tuple<CasExpr, CasExpr> UpdateForDim(float p1, float c, float p2, float length, Dictionary<string, SnapCacheEntry> cachedPos)
        {
            int count = 0;
            if (cachedPos.ContainsKey(FloatToKey(p1))) count++;
            if (cachedPos.ContainsKey(FloatToKey(c))) count++;
            if (cachedPos.ContainsKey(FloatToKey(p2))) count++;
            if (CachedLengths.ContainsKey(FloatToKey(length))) count++;

            CasExpr p1Expr;
            if (cachedPos.ContainsKey(FloatToKey(p1)))
            {
                p1Expr = cachedPos[FloatToKey(p1)].Expr;
            }
            else
            {
                var p1Var = CAS.NewVar();
                Variables[p1Var] = new VariableCacheEntry(p1Var, p1);
                p1Expr = CAS.VarExpr(p1Var);
            }

            CasExpr lengthExpr;
            if (CachedLengths.ContainsKey(FloatToKey(length)))
            {
                lengthExpr = CachedLengths[FloatToKey(length)].Expr;
            }
            else
            {
                var lengthVar = CAS.NewVar();
                Variables[lengthVar] = new VariableCacheEntry(lengthVar, length);
                lengthExpr = CAS.VarExpr(lengthVar);
            }

            var p2Expr = CAS.Add(p1Expr, lengthExpr);
            var cExpr = CAS.Add(p1Expr, CAS.Div(lengthExpr, CAS.ConstExpr(2)));
            AddOrMergeCacheEntry(p1Expr, cachedPos, FloatToKey(p1));
            AddOrMergeCacheEntry(lengthExpr, CachedLengths, FloatToKey(length));
            AddOrMergeCacheEntry(p2Expr, cachedPos, FloatToKey(p2));
            AddOrMergeCacheEntry(cExpr, cachedPos, FloatToKey(c));
            return new Tuple<CasExpr, CasExpr>(p1Expr, lengthExpr);
        }

        public void AddOrMergeCacheEntry(CasExpr expr, Dictionary<string, SnapCacheEntry> cache, string key)
        {
            if (!cache.ContainsKey(key))
            {
                SnapCacheEntry entry = new SnapCacheEntry(expr, cache, key);
                cache[key] = entry;
                CacheEntries.Add(entry);
            }
            else
            {
                var currentExpr = cache[key].Expr;
                var diffExpr = CAS.Sub(expr, currentExpr);
                var vars = new List<CasVar>(CAS.UsedVariables(diffExpr));
                vars.Sort((CasVar var1, CasVar var2) => Variables[var2].Value.CompareTo(Variables[var1].Value));
                foreach (var v in vars)
                {
                    var solution = CAS.MakeZero(diffExpr, v);
                    if (solution != null)
                    {
                        ReplaceVariable(v, solution);
                        return;
                    }
                }
                if (diffExpr.Poly1.Terms.Count != 0)
                {
                    throw new Exception("Contradictory expression");
                }
            }
        }

        public void ReplaceVariable(CasVar variable, CasExpr expr)
        {
            foreach (var entry in CacheEntries)
            {
                entry.Expr = CAS.Substitute(expr, variable, entry.Expr);
            }
        }

        public string FloatToKey(float value)
        {
            return value.ToString("0.00");
        }

        public void RemoveUnusedVariables(List<SlideObject> slideObjects)
        {
            foreach (var entry in CacheEntries)
            {
                entry.InUse = false;
            }
            foreach (var obj in slideObjects)
            {
                FlagIfPresent(CachedXs, obj.X1);
                FlagIfPresent(CachedXs, obj.X2);
                FlagIfPresent(CachedXs, obj.CX);
                FlagIfPresent(CachedYs, obj.Y1);
                FlagIfPresent(CachedYs, obj.Y2);
                FlagIfPresent(CachedYs, obj.CY);
                FlagIfPresent(CachedLengths, obj.Width);
            }
            foreach (var entry in Variables.Values)
            {
                entry.InUse = false;
            }
            foreach (var entry in CacheEntries)
            {
                if(entry.InUse)
                {
                    foreach (var variable in CAS.UsedVariables(entry.Expr))
                    {
                        Variables[variable].InUse = true;
                    }
                }
            }
            foreach (var entry in new List<VariableCacheEntry>(Variables.Values))
            {
                // Remove unused variables
                if (!entry.InUse)
                {
                    Variables.Remove(entry.Variable);
                }
            }
            foreach (var entry in new List<SnapCacheEntry>(CacheEntries))
            {
                // Remove unused cache entries
                foreach (var variable in CAS.UsedVariables(entry.Expr))
                {
                    if (!Variables.ContainsKey(variable))
                    {
                        CacheEntries.Remove(entry);
                        entry.Cache.Remove(entry.Key);
                    }
                }
            }
        }

        public void FlagIfPresent(Dictionary<string, SnapCacheEntry> cache, float value)
        {
            string key = FloatToKey(value);
            if (cache.ContainsKey(key))
            {
                cache[key].InUse = true;
            }
        }
    }

    class SnapCacheEntry
    {
        public CasExpr Expr;
        public bool InUse = false;
        public Dictionary<string, SnapCacheEntry> Cache;
        public string Key;
        public SnapCacheEntry(CasExpr expr, Dictionary<string, SnapCacheEntry> cache, string key)
        {
            Expr = expr;
            Cache = cache;
            Key = key;
        }
    }

    class VariableCacheEntry
    {
        public CasVar Variable;
        public bool InUse = false;
        public float Value;
        public VariableCacheEntry(CasVar variable, float value)
        {
            Variable = variable;
            Value = value;
        }
    }
}
