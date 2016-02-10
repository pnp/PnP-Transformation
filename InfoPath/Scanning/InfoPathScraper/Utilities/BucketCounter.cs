using System;
using System.Collections.Generic;

namespace InfoPathScraper.Utilities
{
	/// <summary>
	/// Utility class that keeps a count of how many occurences of 
	/// a particular key there were. This can be generalized by making the
	/// BucketCounter be templated to the keyType.
	/// </summary>
	class BucketCounter
	{
		private Dictionary<string, Int32> _dictionary;

		public BucketCounter()
		{
			_dictionary = new Dictionary<string, int>();
		}

		/// <summary>
		/// Creates a key and initializes it to a count of zero.
		/// </summary>
		/// <param name="key"></param>
		public void DefineKey(string key)
		{
			if (!_dictionary.ContainsKey(key))
				_dictionary.Add(key, 0);
		}

		/// <summary>
		/// Make sure we have a counter for the key and increment it. 
		/// Note that this is why I use Int32 objects instead of int
		/// </summary>
		/// <param name="key"></param>
		public void IncrementKey(string key)
		{
			DefineKey(key);
			_dictionary[key]++;
		}

		/// <summary>
		/// For each KeyValuePair, return it. I create new ones here so that 
		/// a caller can't accidentally mess up the values in the Dictionary
		/// The buckets are returned in sorted-by-key order because that's more useful than a random order
		/// </summary>
		public IEnumerable<KeyValuePair<string, int>> Buckets
		{
			get
			{
				List<string> orderedKeys = new List<string>();
				foreach (string key in _dictionary.Keys)
					orderedKeys.Add(key);
				orderedKeys.Sort();

				foreach(string key in orderedKeys)
					yield return new KeyValuePair<string, int>(key, _dictionary[key]);
				yield break;
			}
		}
	}
}
