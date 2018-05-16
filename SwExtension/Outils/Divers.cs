using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Outils
{
    public static class DictionaryListe
    {
        public static Boolean AddIfNotExist<TKey, TValue>(this Dictionary<TKey, TValue> dic, TKey key, TValue value)
        {
            if (!dic.ContainsKey(key))
            {
                dic.Add(key, value);
                return true;
            }

            return false;
        }

        public static Boolean AddIfNotExist<T>(this List<T> list, T value)
        {
            if (!list.Contains(value))
            {
                list.Add(value);
                return true;
            }

            return false;
        }

        public static T Last<T>(this List<T> list)
        {
            return list[list.Count - 1];
        }

        public static KeyValuePair<TKey, TValue> KeyValuePair<TKey, TValue>(this Dictionary<TKey, TValue> dictionary, TKey key)
        {
            return new KeyValuePair<TKey, TValue>(key, dictionary[key]);
        }

        /// <summary>
        /// Ajoute la clé ou un à la valeur
        /// </summary>
        /// <typeparam name="TKey"></typeparam>
        /// <param name="dic"></param>
        /// <param name="key"></param>
        /// <returns></returns>
        public static Boolean AddIfNotExistOrPlus<TKey>(this Dictionary<TKey, int> dic, TKey key)
        {
            if (!dic.ContainsKey(key))
            {
                dic.Add(key, 1);
                return true;
            }
            else
                dic[key]++;

            return false;
        }

        /// <summary>
        /// Ajoute 1 à la valeur si la clé existe
        /// </summary>
        /// <typeparam name="TKey"></typeparam>
        /// <param name="dic"></param>
        /// <param name="key"></param>
        /// <returns></returns>
        public static Boolean Plus<TKey>(this Dictionary<TKey, int> dic, TKey key)
        {
            if (dic.ContainsKey(key))
            {
                dic[key]++;
                return true;
            }

            return false;
        }

        /// <summary>
        /// Ajoute la clé si elle n'existe pas et initialise la valeur à 1
        /// </summary>
        /// <typeparam name="TKey"></typeparam>
        /// <param name="dic"></param>
        /// <param name="key"></param>
        /// <returns></returns>
        public static Boolean Ajouter<TKey>(this Dictionary<TKey, int> dic, TKey key)
        {
            if (!dic.ContainsKey(key))
            {
                dic.Add(key, 1);
                return true;
            }

            return false;
        }

        public static void Multiplier<TKey>(this Dictionary<TKey, int> dic, int m)
        {
            foreach (var k in dic.Keys.ToList())
            {
                int v = dic[k];
                dic[k] = v * m;
            }
        }
    }
}
