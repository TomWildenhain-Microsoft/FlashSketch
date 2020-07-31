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
        Dictionary<int, SlideObject> ObjectsToKeepUpdated = null;

        CasSystem CAS;
        public SnapDetector()
        {
            CAS = CasSystem.Instance;
        }

        public void ResizeShape(Dictionary<int, SlideObject> slideObjects, int shapeId, float width, float height)
        {
            UpdateSnapCache(slideObjects);
            FindExprsForSlideObjs(slideObjects);
            ObjectsToKeepUpdated = slideObjects;
            var widthExpr = MakeExprForFloat(width, CachedLengths);
            EqualizeExpressions(widthExpr, slideObjects[shapeId].WidthExpr, widthExpr, slideObjects);
            var heightExpr = MakeExprForFloat(height, CachedLengths);
            EqualizeExpressions(heightExpr, slideObjects[shapeId].HeightExpr, heightExpr, slideObjects);
            ObjectsToKeepUpdated = null;
        }

        public void EqualizeLongerDims(Dictionary<int, SlideObject> slideObjects, int shapeId1, int shapeId2)
        {
            UpdateSnapCache(slideObjects);
            FindExprsForSlideObjs(slideObjects);
            CasExpr expr1 = slideObjects[shapeId1].LongerDimExpr();
            CasExpr expr2 = slideObjects[shapeId2].LongerDimExpr();
            EqualizeExpressions(expr1, expr2, expr1, slideObjects);
        }

        public void EqualizeHeights(Dictionary<int, SlideObject> slideObjects, int shapeId1, int shapeId2)
        {
            UpdateSnapCache(slideObjects);
            FindExprsForSlideObjs(slideObjects);
            SlideObject shape1 = slideObjects[shapeId1];
            SlideObject shape2 = slideObjects[shapeId2];
            EqualizeExpressions(shape1.HeightExpr, shape2.HeightExpr, shape1.HeightExpr, slideObjects);
        }

        public void EqualizeWidths(Dictionary<int, SlideObject> slideObjects, int shapeId1, int shapeId2)
        {
            UpdateSnapCache(slideObjects);
            FindExprsForSlideObjs(slideObjects);
            SlideObject shape1 = slideObjects[shapeId1];
            SlideObject shape2 = slideObjects[shapeId2];
            EqualizeExpressions(shape1.WidthExpr, shape2.WidthExpr, shape1.WidthExpr, slideObjects);
        }

        public void MakeSquare(Dictionary<int, SlideObject> slideObjects, int shapeId)
        {
            UpdateSnapCache(slideObjects);
            FindExprsForSlideObjs(slideObjects);
            SlideObject shape = slideObjects[shapeId];
            EqualizeExpressions(shape.HeightExpr, shape.WidthExpr, shape.ShorterDimExpr(), slideObjects);
        }

        public void SetShapeWidth(Dictionary<int, SlideObject> slideObjects, int shapeId, float width)
        {
            UpdateSnapCache(slideObjects);
            SlideObject shape = slideObjects[shapeId];
            var widthExpr = MakeExprForFloat(width, CachedLengths);
            AddOrMergeCacheEntry(widthExpr, CachedLengths, FloatToKey(width));
            widthExpr = CachedLengths[FloatToKey(width)].Expr;
            FindExprsForSlideObjs(slideObjects);
            EqualizeExpressions(widthExpr, shape.WidthExpr, widthExpr, slideObjects);
        }

        public void SetShapeHeight(Dictionary<int, SlideObject> slideObjects, int shapeId, float height)
        {
            UpdateSnapCache(slideObjects);
            SlideObject shape = slideObjects[shapeId];
            var heightExpr = MakeExprForFloat(height, CachedLengths);
            AddOrMergeCacheEntry(heightExpr, CachedLengths, FloatToKey(height));
            heightExpr = CachedLengths[FloatToKey(height)].Expr;
            FindExprsForSlideObjs(slideObjects);
            EqualizeExpressions(heightExpr, shape.HeightExpr, heightExpr, slideObjects);
        }

        public void EqualizeExpressions(CasExpr expr1, CasExpr expr2, CasExpr invariant, Dictionary<int, SlideObject> slideObjects)
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
                ObjectsToKeepUpdated = slideObjects;
                ReplaceVariable(bestVar, bestSolution);
                ObjectsToKeepUpdated = null;
            }

            // FindExprsForSlideObjs(slideObjects);
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

        public void FindExprsForSlideObjs(Dictionary<int, SlideObject> slideObjects)
        {
            foreach (var obj in slideObjects.Values)
            {
                obj.XExpr = CachedXs[FloatToKey(obj.X1)].Expr;
                obj.WidthExpr = CachedLengths[FloatToKey(obj.Width)].Expr;
                obj.YExpr = CachedYs[FloatToKey(obj.Y1)].Expr;
                obj.HeightExpr = CachedLengths[FloatToKey(obj.Height)].Expr;
            }
        }

        public void UpdateSnapCacheAfterHDist(Dictionary<int, SlideObject> slideObjects, List<Tuple<int, float>> ids)
        {
            foreach (var obj in slideObjects.Values)
            {
                UpdateForDim(obj.X1, obj.CX, obj.X2, obj.Width, CachedXs);
                UpdateForDim(obj.Y1, obj.CY, obj.Y2, obj.Height, CachedYs);
            }
            FindExprsForSlideObjs(slideObjects);
            for (int i = 0; i < ids.Count - 1; i++)
            {
                var obj1 = slideObjects[ids[i].Item1];
                var obj2 = slideObjects[ids[i + 1].Item1];
                var obj1X2 = CAS.Add(obj1.XExpr, obj1.WidthExpr);
                var obj2X1 = obj2.XExpr;
                var dist = CAS.Sub(obj2X1, obj1X2);
                var key = FloatToKey(obj2.X1 - obj1.X2);
                AddOrMergeCacheEntry(dist, CachedLengths, key);
                FindExprsForSlideObjs(slideObjects);
            }
            RemoveUnusedVariables(slideObjects);
        }

        public void PrintConstraints(Dictionary<int, SlideObject> slideObjects)
        {
            FindExprsForSlideObjs(slideObjects);
            FlashSketch.Instance.ClearNotes();
            FlashSketch.Instance.PrintToNotes(" --- Constraints --- ");
            foreach (var obj in slideObjects.Values)
            {
                string line = obj.ShapeName + ": {width = \"" + obj.WidthExpr.ToString() + "\", ";
                line += "height = \"" + obj.HeightExpr.ToString() + "\", ";
                line += "x = \"" + obj.XExpr.ToString() + "\", ";
                line += "y = \"" + obj.YExpr.ToString() + "\"}";
                FlashSketch.Instance.PrintToNotes(line);
            }
            int i = 0;
            string res = "";
            foreach (var v in Variables.Values)
            {
                if (v.InUse)
                {
                    res += "\n" + (v.Variable.Name + " = " + v.Value);
                    i++;
                }
            }
            FlashSketch.Instance.PrintToNotes(" --- Variables (Count=" + i + ") --- " + res);
        }

        public void UpdateSnapCacheAfterVDist(Dictionary<int, SlideObject> slideObjects, List<Tuple<int, float>> ids)
        {
            foreach (var obj in slideObjects.Values)
            {
                UpdateForDim(obj.X1, obj.CX, obj.X2, obj.Width, CachedXs);
                UpdateForDim(obj.Y1, obj.CY, obj.Y2, obj.Height, CachedYs);
            }
            FindExprsForSlideObjs(slideObjects);
            for (int i = 0; i < ids.Count - 1; i++)
            {
                var obj1 = slideObjects[ids[i].Item1];
                var obj2 = slideObjects[ids[i + 1].Item1];
                var obj1Y2 = CAS.Add(obj1.YExpr, obj1.HeightExpr);
                var obj2Y1 = obj2.YExpr;
                var dist = CAS.Sub(obj2Y1, obj1Y2);
                var key = FloatToKey(obj2.Y1 - obj1.Y2);
                AddOrMergeCacheEntry(dist, CachedLengths, key);
                FindExprsForSlideObjs(slideObjects);
            }
            RemoveUnusedVariables(slideObjects);
            // PrintConstraints(slideObjects);
        }

        public void UpdateSnapCache(Dictionary<int, SlideObject> slideObjects)
        {
            foreach (var obj in slideObjects.Values)
            {
                UpdateForDim(obj.X1, obj.CX, obj.X2, obj.Width, CachedXs);
                UpdateForDim(obj.Y1, obj.CY, obj.Y2, obj.Height, CachedYs);
            }
            RemoveUnusedVariables(slideObjects);
            // PrintConstraints(slideObjects);
        }

        public CasExpr MakeExprForFloat(float value, Dictionary<string, SnapCacheEntry> cache)
        {
            string key = FloatToKey(value);
            if (cache.ContainsKey(key) && CAS.UsedVariables(cache[key].Expr).Count <= 1)
            {
                return cache[key].Expr;
            }
            else
            {
                var v = CAS.NewVar();
                Variables[v] = new VariableCacheEntry(v, value);
                return CAS.VarExpr(v);
            }
        }

        public void UpdateForDim(float p1, float c, float p2, float length, Dictionary<string, SnapCacheEntry> cachedPos)
        {

            CasExpr p1Expr = MakeExprForFloat(p1, cachedPos);
            AddOrMergeCacheEntry(p1Expr, cachedPos, FloatToKey(p1));
            p1Expr = cachedPos[FloatToKey(p1)].Expr;

            CasExpr lengthExpr = MakeExprForFloat(length, CachedLengths);
            AddOrMergeCacheEntry(lengthExpr, CachedLengths, FloatToKey(length));
            lengthExpr = CachedLengths[FloatToKey(length)].Expr;

            var p2Expr = CAS.Add(p1Expr, lengthExpr);
            AddOrMergeCacheEntry(p2Expr, cachedPos, FloatToKey(p2));

            var cExpr = CAS.Add(p1Expr, CAS.Div(lengthExpr, CAS.ConstExpr(2)));
            AddOrMergeCacheEntry(cExpr, cachedPos, FloatToKey(c));
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
            if (ObjectsToKeepUpdated != null)
            {
                foreach (var obj in ObjectsToKeepUpdated.Values)
                {
                    obj.HeightExpr = CAS.Substitute(expr, variable, obj.HeightExpr);
                    obj.WidthExpr = CAS.Substitute(expr, variable, obj.WidthExpr);
                    obj.XExpr = CAS.Substitute(expr, variable, obj.XExpr);
                    obj.YExpr = CAS.Substitute(expr, variable, obj.YExpr);
                }
            }
        }

        public string FloatToKey(float value)
        {
            return value.ToString("0.00");
        }

        public void RemoveUnusedVariables(Dictionary<int, SlideObject> slideObjects)
        {
            foreach (var entry in CacheEntries)
            {
                entry.InUse = false;
            }
            foreach (var obj in slideObjects.Values)
            {
                FlagIfPresent(CachedXs, obj.X1);
                FlagIfPresent(CachedXs, obj.X2);
                FlagIfPresent(CachedXs, obj.CX);
                FlagIfPresent(CachedYs, obj.Y1);
                FlagIfPresent(CachedYs, obj.Y2);
                FlagIfPresent(CachedYs, obj.CY);
                FlagIfPresent(CachedLengths, obj.Width);
                FlagIfPresent(CachedLengths, obj.Height);
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
            return;
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
