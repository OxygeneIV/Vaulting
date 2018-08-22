using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using Framework.Utils;
using NLog;
using OpenQA.Selenium;

namespace Framework.PageObjects
{
    public static class PageObjectFactory
    {
        internal static Logger Log = LogManager.GetCurrentClassLogger();

        /// <summary>
        /// Create a PageObject of type T (optionally setting initial element)
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="parent"></param>
        /// <param name="element"></param>
        /// <returns></returns>
        public static T CreatePageObject<T>(IParent parent, IWebElement element = null) where T : PageObject, new()
        {
            var po = Init<T>(parent);
            po.SetWrappedElement(element);
            return po;
        }

        /// <summary>
        /// Convert to a PageObject of type T (optionally setting initial element)
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        public static T ConvertPageObject<T>(PageObject original) where T : PageObject, new()
        {
            var po = Init<T>(original.Parent);
            po.Locator = original.Locator;
            po.SetWrappedElement(original.CachedElement);
            return po;
        }

        /// <summary>
        /// Get an element from an IEnumerable<list type="T"></list> as a PageObject, setting a "Filter"-predicate to find/re-find it
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="list"></param>
        /// <param name="predicate"></param>
        /// <returns></returns>
        public static T GetElement<T>(this IEnumerable<T> list, Func<T, bool> predicate) where T : PageObject, new()
        {
            Log.Info($"Registering FindMeCandidates() for PageObject of type {typeof(T)}");
            var po = Init<T>(((PageObjectListProxy<T>)list).Parent);
            Log.Info($"Inheriting {((PageObjectListProxy<T>)list).PageObjectFilter.Count} filters");
            var newFilterList =  new List<Func<T,bool>>(); 
            newFilterList.AddRange(((PageObjectListProxy<T>)list).PageObjectFilter);
            newFilterList.Add(predicate);
            po.FindMeCandidates = () => po.FindMeCandidates_(newFilterList);
            Log.Info($"Registering FindMeCandidates() completed for PageObject of type {typeof(T)} ");
            return po;
        }

        public static U GetElementAs<U,T>(this IEnumerable<T> inlist, Func<T, bool> predicate)
            where T : PageObject, new() where U : T, new()
        {
            Log.Info($"Registering FindMeCandidates() for PageObject of type {typeof(T)}");
            var list = new PageObjectListProxy<U>(((PageObjectListProxy<T>) inlist).Parent,((PageObjectListProxy<T>) inlist).Locator);
            var po = Init<U>(list.Parent);
            Log.Info($"Inheriting {list.PageObjectFilter.Count} filters");
            var newFilterList = new List<Func<U, bool>>();
            newFilterList.AddRange(list.PageObjectFilter);
            newFilterList.Add(predicate);
            po.FindMeCandidates = () => po.FindMeCandidates_(newFilterList);
            Log.Info($"Registering FindMeCandidates() completed for PageObject of type {typeof(U)} ");
            return po;
        }


        /// <summary>
        /// Get an element from an IEnumerable<list type="T"></list> as a PageObject, setting a locatorAttribute via the parameter
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="list"></param>
        /// <param name="locatorAttributeParam"></param>
        /// <returns></returns>
        public static T GetElement<T>(this IEnumerable<T> list, string locatorAttributeParam) where T : PageObject, new()
        {
            Log.Info($"Registering FindMeCandidates() for PageObject of type {typeof(T)} with locatorAttribute functionality");
            var po = CreatePageObject<T>(((PageObjectListProxy<T>) list).Parent);
            po.Locator = po.LocatorFunc(locatorAttributeParam);

            Log.Info($"Inheriting {((PageObjectListProxy<T>)list).PageObjectFilter.Count} filters");
            var newFilterList = new List<Func<T, bool>>();
            newFilterList.AddRange(((PageObjectListProxy<T>)list).PageObjectFilter);
            po.FindMeCandidates = () => po.FindMeCandidates_(newFilterList);
            Log.Info($"Registering FindMeCandidates() completed for PageObject of type {typeof(T)} ");
            return po;
        }




        /// <summary>
        /// Get an element from an IEnumerable<list type="T"></list> as a PageObject, setting a locatorAttribute via the parameter
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="list"></param>
        /// <param name="itemLocatorAttributeParam"></param>
        /// <returns></returns>
        public static T Item<T>(this IEnumerable<T> list, params object[] itemLocatorAttributeParam) where T : PageObject, new()
        {
            Log.Info($"Registering FindMeCandidates() for PageObject of type {typeof(T)} with locatorAttribute functionality");
            var poList = (PageObjectListProxy<T>) list;

            var po = CreatePageObject<T>(poList.Parent);

            var formatString = poList.ItemLocator.Formatter;
            po.Locator = new LocatorAttribute(poList.ItemLocator.How, string.Format(formatString, itemLocatorAttributeParam));

            Log.Info($"Inheriting {poList.PageObjectFilter.Count} filters");
            var newFilterList = new List<Func<T, bool>>();
            newFilterList.AddRange(poList.PageObjectFilter);
            po.FindMeCandidates = () => po.FindMeCandidates_(newFilterList);
            Log.Info($"Registering FindMeCandidates() completed for PageObject of type {typeof(T)} ");
            return po;
        }


        public static void AddDictionaryItem<T>(this IEnumerable<T> list, Enum key, string value) where T : PageObject, new()
        {
            ((PageObjectListProxy<T>)list).ItemDictionary.Add(key,value);
        }

        public static T GetElement<T>(this IEnumerable<T> list, Enum dictionaryItem) where T : PageObject, new()
        {
            Log.Info($"Registering FindMeCandidates() for PageObject of type {typeof(T)} with locatorAttribute functionality");
            var po = CreatePageObject<T>(((PageObjectListProxy<T>)list).Parent);

            var enummapperExists = ((PageObjectListProxy<T>)list).ItemDictionary.ContainsKey(dictionaryItem);
            if (!enummapperExists)
            {
                Log.Fatal($"Enum {dictionaryItem} does not exist in the ItemDictionary");
                throw new Exception($"Missing item dictionary {dictionaryItem} in List < {nameof(T)} > ");
            }

            var poString = ((PageObjectListProxy<T>)list).ItemDictionary[dictionaryItem];
            var formatString = ((PageObjectListProxy<T>) list).ItemLocator.Formatter;

            var newLoc = new LocatorAttribute(((PageObjectListProxy<T>)list).ItemLocator.How,string.Format(formatString,poString));
            po.Locator = newLoc;

            Log.Info($"Inheriting {((PageObjectListProxy<T>)list).PageObjectFilter.Count} filters");
            var newFilterList = new List<Func<T, bool>>();
            newFilterList.AddRange(((PageObjectListProxy<T>)list).PageObjectFilter);
            po.FindMeCandidates = () => po.FindMeCandidates_(newFilterList);
            Log.Info($"Registering FindMeCandidates() completed for PageObject of type {typeof(T)} ");
            return po;
        }


        public static PageObjectListProxy<T> GetElements<T>(this IEnumerable<T> list, Func<T, bool> predicate) where T : PageObject, new()
        {
            Log.Info($"Registering FindMeCandidates() for IEnumerable<T> of type {typeof(T)} item");
            var parent = ((PageObjectListProxy<T>) list).Parent;
            var locator = ((PageObjectListProxy<T>) list).Locator;

            // Inherit and add another filter
            var newFilterList = new List<Func<T, bool>>();
            newFilterList.AddRange(((PageObjectListProxy<T>)list).PageObjectFilter);
            newFilterList.Add(predicate);

            var theList = new PageObjectListProxy<T>(parent, locator)
            {
                PageObjectFilter = newFilterList
            };
            Log.Info($"Registering FindMeCandidates() completed for PageObject of type {typeof(T)} ");
            return theList;
        }

        /// <summary>
        ///  Get an element from an IEnumerable<list type="T"></list> as a PageObject, find by Index
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="list"></param>
        /// <param name="index"></param>
        /// <returns></returns>
        public static T GetElement<T>(this IEnumerable<T> list, int index) where T : PageObject, new()
        {
            Log.Info($"Registering FindMeCandidates() for PageObject of type {typeof(T)} item with index {index}");

            var parent = ((PageObjectListProxy<T>)list).Parent;

            var po = Init<T>(parent);

            // Inherit filters from list
            var newFilterList = new List<Func<T, bool>>();

            newFilterList.AddRange(((PageObjectListProxy<T>)list).PageObjectFilter);

            po.FindMeCandidates = () => po.FindMeCandidates_(newFilterList, index);
            Log.Info($"Registering FindMeCandidates() completed for PageObject item with index {index}");
            return po;
        }

        /// <summary>
        /// Create a PageObject of type T (optionally setting a specific locator)
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="parent"></param>
        /// <param name="locator"></param>
        /// <returns></returns>
        public static T CreatePageObject<T>(IParent parent, LocatorAttribute locator) where T : PageObject, new()
        {
            var po = Init<T>(parent);
            po.Locator = locator;
            return po;
        }

        public static T SwitchToMe<T>(this T pageObject) where T : PageObject, new()
        {
            var myContext = pageObject.WindowHandle;

            if (myContext == null)
            {
               Log.Warn($"Can only switch to pages with WindowHandle set, {nameof(T)} object is missing it");    
            }

            Log.Info($"Switching to Driver context {pageObject.WindowHandle}");
            pageObject.WebDriver.SwitchToWindow(myContext);
            Log.Info($"Returning {pageObject}");
            return pageObject;
        }

        /// <summary>
        /// Initialize a PageObject, instantiating all PageObject children using reflection
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="parent"></param>
        /// <returns></returns>
        public static T Init<T>(IParent parent) where T : PageObject, new()
        {
            var obj = new T
            {
                Parent = parent
               ,WindowHandle = parent.WindowHandle
            };
            var type = obj.GetType();

            var locatorAttribute = (LocatorAttribute)Attribute.GetCustomAttribute(type, typeof(LocatorAttribute));
            obj.Locator = locatorAttribute ?? throw new Exception($"No locator found for {typeof(T)}, need one !");            

            // Analyze children
            Init(obj);

            return obj;
        }

        /// <summary>
        /// Analyze <param name="pageObject"></param> and discover PageObject children
        /// </summary>
        /// <param name="pageObject"></param>
        public static void Init(PageObject pageObject)
        {
            // Call Translation
            if (TranslationManager.Active)
            {
                pageObject.TranslateRegistration();
            }

            var memberInfoList = FetchPageObjectMembers_(pageObject);

            foreach (var mInfo in memberInfoList)
            {


                if (mInfo.GetCustomAttribute(typeof(CompilerGeneratedAttribute)) != null)
                    continue;

                LocatorAttribute locatorAttribute;

                // Field or Property
                var thisType = mInfo.MemberType;

                var myType = (thisType == MemberTypes.Field)
                                ? ((FieldInfo) mInfo).FieldType
                                : ((PropertyInfo) mInfo).PropertyType;

                // ignore properties without setters, they are probably just helper properties
                if (mInfo.MemberType == MemberTypes.Property && !((PropertyInfo) mInfo).CanWrite)
                    continue;

                var iList = false;

                // A list
                if (myType.IsGenericType && myType.GetGenericTypeDefinition() == typeof(IEnumerable<>))
                {
                    var tmpType = myType.GetGenericArguments()[0];
                    if (tmpType.IsSubclassOf(typeof(PageObject)))
                    {
                        locatorAttribute = (LocatorAttribute)Attribute.GetCustomAttribute(tmpType, typeof(LocatorAttribute));
                        iList = true;
                    }
                    else
                    {
                        continue;
                    }
                }
                else
                {
                    locatorAttribute = (LocatorAttribute)Attribute.GetCustomAttribute(myType, typeof(LocatorAttribute));
                }

                // Check if overridden
                locatorAttribute = (LocatorAttribute)Attribute.GetCustomAttribute(mInfo, typeof(LocatorAttribute))
                                   ?? locatorAttribute;

                if (locatorAttribute == null)
                {
                    throw new Exception($"No locator found for {myType}, need one !");
                }

                object o;

                if (iList)
                {
                    var listType = typeof(PageObjectListProxy<>);
                    var tmpType = myType.GetGenericArguments()[0];
                    var concreteType = listType.MakeGenericType(tmpType);
                    o = Activator.CreateInstance(concreteType, pageObject, locatorAttribute);
                    var itemLocatorAttribute = (ItemLocatorAttribute)Attribute.GetCustomAttribute(mInfo, typeof(ItemLocatorAttribute));
                    if (itemLocatorAttribute != null)
                    {
                        SetPropertyValue_(o, "ItemLocator", itemLocatorAttribute);
                    }

                    var dictionaryAttributes = Attribute.GetCustomAttributes(mInfo, typeof(DictionaryItemAttribute));

                    if (dictionaryAttributes.Any())
                    {
                        var itemDictionary = new Dictionary<Enum, string>();

                        foreach (var dictionaryAttribute in dictionaryAttributes)
                        {
                            var myDictionaryItem = (DictionaryItemAttribute) dictionaryAttribute;
                            itemDictionary.Add(myDictionaryItem.ItemKey, myDictionaryItem.Value);
                        }
                        SetPropertyValue_(o, "ItemDictionary", itemDictionary);
                    }
                }
                else
                {

                    o = Activator.CreateInstance(myType);
                    SetPropertyValue_(o, "Parent", pageObject);
                    SetPropertyValue_(o, "Locator", locatorAttribute);
                }

                switch (thisType)
                {
                    case MemberTypes.Field:
                        var fieldInfo = mInfo as FieldInfo;
                        fieldInfo?.SetValue(pageObject, o);
                        break;
                    case MemberTypes.Property:
                        var propertyInfo = mInfo as PropertyInfo;
                        propertyInfo?.SetValue(pageObject, o, null);
                        break;
                    default:
                        throw new Exception("TYPE MISMATCH ");
                }



                // No need to do recursive things if a PageObject Collection was found
                // Each Item will be do its own Init when accessed
                if (!IsSubclassOfRawGeneric_(typeof(IEnumerable<>), myType))
                {
                    Init((PageObject)o);
                }
            }
        }


        private static bool IsSubclassOfRawGeneric_(Type generic, Type toCheck)
        {
            while (toCheck != null && toCheck != typeof(object))
            {
                var cur = toCheck.IsGenericType ? toCheck.GetGenericTypeDefinition() : toCheck;
                if (generic == cur)
                {
                    return true;
                }

                toCheck = toCheck.BaseType;
            }

            return false;
        }

        private static List<MemberInfo> FetchPageObjectMembers_<T>(T myObject)
        {
            var type = myObject.GetType();
            const BindingFlags bindingFlags =
              BindingFlags.Instance | BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic
              | BindingFlags.FlattenHierarchy;
              

            // Fetch all fields and properties that are PageObjects
            var members = new List<MemberInfo>();
            members.AddRange(type.GetFields(bindingFlags).Where(p => p.FieldType.IsSubclassOf(typeof(PageObject))));
            members.AddRange(type.GetProperties(bindingFlags).Where(p => p.PropertyType.IsSubclassOf(typeof(PageObject))));
            members.AddRange(type.GetFields(bindingFlags).Where(p => IsSubclassOfRawGeneric_(typeof(IEnumerable<>), p.FieldType)));
            members.AddRange(type.GetProperties(bindingFlags).Where(p => IsSubclassOfRawGeneric_(typeof(IEnumerable<>), p.PropertyType)));
            return members;
        }

        private static PropertyInfo GetPropertyInfo_(Type type, string propertyName)
        {
            PropertyInfo propInfo;
            do
            {
                propInfo = type.GetProperty(propertyName,
                       BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
                type = type.BaseType;
            }
            while (propInfo == null && type != null);
            return propInfo;
        }

        private static void SetPropertyValue_(this object obj, string propertyName, object val)
        {
            if (obj == null)
                throw new ArgumentNullException("obj");
            var objType = obj.GetType();
            var propInfo = GetPropertyInfo_(objType, propertyName);
            if (propInfo == null)
                throw new ArgumentOutOfRangeException(nameof(propertyName),
                    $"Couldn't find property {propertyName} in type {objType.FullName}");
            propInfo.SetValue(obj, val, null);
        }
    }
}
