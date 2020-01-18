using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpreadsheetUtil
{
	//??? change the name of this file
	public interface IRefCount
	{
		int GetRefCount();
		void IncrementRefCount();
		void DecrementRefCount();
	}

	public interface IGetIntegerID
	{
		int GetGetID();
	}

	internal sealed class ImmutableRefCountSet<T> where T : IRefCount, IGetIntegerID// where T : struct - maybe T needs to implement some interfaces for decrementing/incrementing ref count and getting ref count and also ???
	{
		private const int MaxListElems = 16;
		private List<T> uniqueList = new List<T>(MaxListElems);
		private HashSet<T> set = null; // create this if needed when the List would get too large

		// return true if newItem is found and not added (ref count is incrememented), false if not found and added (ref count set to 1)
		public bool AddOrFindItem(T newItem)
		{
			bool isFound = false;
			if (set == null)
			{
				if (uniqueList.Count < MaxListElems)
				{
					
				}
				else
				{
					set = new HashSet<T>(uniqueList);
				}
			}
			
			if (set != null)
			{
				// TryGetValue requires .NET 4.7.2
				if (set.TryGetValue(newItem, out T actualValue))
				{
					//???
				}
			}
			return isFound;
		}
	}
}
