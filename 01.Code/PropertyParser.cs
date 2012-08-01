using System;
using System.Collections.Generic;
using System.Text;
using System.Reflection;

namespace NPOI.Wrapper
{
    public class PropertyParser
    {
        public static PropertyInfo GetProperty(object objContext, string Name, out bool IsList, out object Context)
        {
            PropertyInfo pObj = null;
            List<string> pNames = ParseName(Name);
            Context = objContext;
            IsList = false;

            for (int i = 0; i < pNames.Count; i++)
            {
                Type tContext = Context.GetType();

                if (pNames[i].StartsWith("[]"))
                {
                    PropertyIndex pIndex = ParseIndex(pNames[i].Substring(2));

                    MemberInfo[] M = tContext.GetDefaultMembers();
                    foreach (MemberInfo m in M)
                    {
                        IsList = true;

                        if (pIndex.Types.Length == 0)
                        {
                            if (i + 1 < pNames.Count)
                            {
                                if (!tContext.IsGenericType)
                                {
                                    throw new Exception(
                                        string.Format("Parse Error On [{0}].[{1}]@{2}", pNames[i], Name, tContext.FullName)
                                        );
                                }
                                Type[] subTypes = tContext.GetGenericArguments();
                                if (subTypes.Length != 1)
                                {
                                    throw new Exception(
                                        string.Format("Parse Error On [{0}].[{1}]@{2}", pNames[i], Name, tContext.FullName)
                                        );
                                }
                                pObj = subTypes[0].GetProperty(pNames[i + 1]);
                            }
                            else
                            {
                                pObj = tContext.GetProperty(m.Name);
                            }
                            return pObj;
                        }

                        pObj = tContext.GetProperty(m.Name, pIndex.Types);
                        if (pObj != null)
                        {
                            if (i < pNames.Count - 1)
                            {
                                Context = pObj.GetValue(Context, pIndex.Value);
                            }
                            break;
                        }
                    }
                }
                else
                {
                    pObj = Context.GetType().GetProperty(pNames[i]);
                    if (i < pNames.Count - 1)
                    {
                        Context = pObj.GetValue(Context, null);
                        if (Context == null)
                        {
                            throw new Exception(
                                string.Format("Property Is Null [{0}]@[{1}]", pNames[i], Context.GetType().FullName)
                                );
                        }
                    }
                }
            }

            return pObj;
        }

        static List<string> ParseName(string Name)
        {
            int MatchCounter = 0;

            StringBuilder NameCache = new StringBuilder();
            List<string> Names = new List<string>();
            for (int i = 0; i < Name.Length; i++)
            {
                if (Name[i] == '.')
                {
                    if (i > 0 && Name[i - 1] == '.')
                    {
                        throw new Exception(string.Format("Parse Error On:{0}@{1}", Name[i], i));
                    }
                    else if (i == Name.Length - 1)
                    {
                        throw new Exception(string.Format("Parse Error On:{0}@{1}", Name[i], i));
                    }
                    else if (NameCache.Length > 0)
                    {
                        Names.Add(NameCache.ToString());
                        NameCache.Remove(0, NameCache.Length);
                    }
                }
                else if (Name[i] == '[')
                {
                    MatchCounter += 1;
                    if (MatchCounter != 1)
                    {
                        throw new Exception(string.Format("Parse Error On:{0}@{1}", Name[i], i));
                    }
                    else if (i == Name.Length - 1)
                    {
                        throw new Exception(string.Format("Parse Error On:{0}@{1}", Name[i], i));
                    }
                    else if (NameCache.Length > 0)
                    {
                        Names.Add(NameCache.ToString());
                        NameCache.Remove(0, NameCache.Length);
                    }
                }
                else if (Name[i] == ']')
                {
                    MatchCounter += 1;
                    if (MatchCounter != 2)
                    {
                        throw new Exception(string.Format("Parse Error On:{0}@{1}", Name[i], i));
                    }
                    else if (Name.Length > i + 1 && Name[i + 1] != '.' && Name[i + 1] != '[')
                    {
                        throw new Exception(string.Format("Parse Error On:{0}@{1}", Name[i], i));
                    }

                    Names.Add("[]" + NameCache.ToString());
                    NameCache.Remove(0, NameCache.Length);

                    MatchCounter = 0;
                }
                else
                {
                    NameCache.Append(Name[i]);
                }
            }

            if (NameCache.Length > 0)
            {
                Names.Add(NameCache.ToString());
            }
            NameCache.Remove(0, NameCache.Length);
            NameCache = null;

            return Names;
        }

        class PropertyIndex
        {
            public Type[] Types { get; set; }

            public object[] Value { get; set; }
        }

        static PropertyIndex ParseIndex(string Index)
        {
            PropertyIndex pIndex = new PropertyIndex();

            List<Type> pTypes = new List<Type>();
            List<object> pValues = new List<object>();

            int indexValue = 0;
            string[] Indexes = Index.Split(',');
            for (int i = 0; i < Indexes.Length; i++)
            {
                if (Indexes[i] == "")
                {
                    continue;
                }
                else if (int.TryParse(Indexes[i], out indexValue))
                {
                    pTypes.Add(typeof(int));
                    pValues.Add(indexValue);
                }
                else if (Indexes[i].StartsWith("'") && Indexes[i].EndsWith("'"))
                {
                    pTypes.Add(typeof(string));
                    pValues.Add(
                        Indexes[i].Substring(1, Indexes[i].Length - 2)
                        );
                }
                else
                {
                    throw new Exception(
                        string.Format("Parse Error On Index[{0}]@[{1}]", Indexes[i], Index)
                        );
                }
            }

            pIndex.Types = pTypes.ToArray();
            pIndex.Value = pValues.ToArray();

            return pIndex;
        }
    }
}
