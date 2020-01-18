using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpreadsheetUtil
{
	public class BasicHashSet<T> : IEnumerable<T>
	{
		private const int InitialItemsTableSize = 16;

		// what if someone wants less than double the # of items allocated next time??? - need to get a prime between 2 numbers in this list? - just let them specify the number, which maybe should be a prime
		static readonly int[] indexArraySizeArray = { 7, 17, 37, 79, 163, 331, 673, 1361, 2729, 5471, 10949, 21911, 43853, 87719, 175447,
			350899, 701819, 1403641, 2807303, 5614657, 11229331, 22458671, 44917381, 89834777, 179669557, 359339171, 718678369, 1437356741, 2147483629 /* don't use int.MaxValue (2147483647), so capacity is allowed up to the largest signed int */ };

		int currentIndexIntoSizeArray;

		// are having private backing variables faster to access than properties???
		// One argument I’ve heard for using fields over properties is that “fields are faster”, but for trivial properties that’s actually not true, as the CLR’s Just-In-Time (JIT) compiler will inline
		// the property access and generate code that’s as efficient as accessing a field directly.
		// But this suggests that the JIT will have to do extra work (cpu time) to do this?

		private bool isHashing; // start out just using the itemsTable as an array without hashing

		private int usedItemsCount;

		private double loadFactor = .75;

		private int usedItemsLoadFactorThreshold;

		private int firstBlankAtEndIndex; // this is needed because if items are removed, they get added into the blank list starting at nodeArray[0].nextIndex, but we may want to TrimExcess capacity, so this is a quick way to see what the ExcessCapacity is

		private int initialArraySize;

		private IEqualityComparer<T> comparer;

		//private int indexArraySize; // this could be the stored size of the array because it is used every time we lookup any item and it might be faster than always getting indexArray.Length
		//private int usedIndexCount; // this would be like a usedItemsCount for the indexArray - this could then be used to determine the # of buckets and the ratio of buckets to items - maybe this is what load is???

		private int[] indexArray; // make the index table a primary number to make the mod function less predictable - use a constant array of prime numbers

		private TNode[] nodeArray;

		private T[] initialArray;

		private int currentNodeIdx; // needed for GetEnumerator

		//private struct NodeLocation
		//{
		//	public NodeLocation(int nodeIndex, int priorNodeIndex, int indexArrayIndex)
		//	{
		//		this.nodeIndex = nodeIndex;
		//		this.priorNodeIndex = priorNodeIndex;
		//		this.indexArrayIndex = indexArrayIndex;
		//	}

		//	public int nodeIndex;
		//	public int priorNodeIndex; // 0 means there was no prior node and instead indexArrayIndex is used
		//	public int indexArrayIndex;
		//}

		private struct TNode
		{
			// putting these variables together (instead of having a separate index array - or an extension of the indexTable array) makes them close in memory, which might make things faster with cpu caching

			public int nextIndex;

			public int hashOrNextIndexForBlanks; // the cached hash code of the item - this is so we don't have to call GetHashCode multiple times, also doubles as a nextIndex for blanks, since blanks don't need a hash code anymore

			public T item;

			public TNode(T elem, int nextIndex, int hash)
			{
				this.item = elem;

				this.nextIndex = nextIndex;

				this.hashOrNextIndexForBlanks = hash;
			}

			public TNode(T elem)
			{
				this.item = elem;

				this.nextIndex = 0; // 0 indicates a blank item

				this.hashOrNextIndexForBlanks = 0;
			}
		}

		// Note: instead of blindly allocating double the # of items each time
		// have a way to specify the # of items that should be allocated next time
		// this is useful if the code knows something, ex. If adding dates from 1950 to 2019 in order we know if we are in 2018 that there can't be many more dates
		// so allocating 2x the amount would be a waste, the program could (at every year of processing, set the allocation to 366 * years remaining to be processed or something like that) -
		// this would be better than 2x if there were alot of dates (or dates and something else being the value of each item)

		public BasicHashSet(int initialCapacity = -1, int initialArraySize = -1, IEqualityComparer<T> comparer = null)
		{
			//??? what about an initialArraySize = 0 means go straight into hashing

			this.comparer = comparer;
			this.initialArraySize = initialArraySize;
			CreateInitialArray(initialArraySize);
		}

		public BasicHashSet(IEnumerable<T> collection, bool areAllCollectionItemsDefinitelyUnique = false, int initialCapacity = -1, int initialArraySize = -1, IEqualityComparer<T> comparer = null)
		{
			//??? what about an initialArraySize = 0 means go straight into hashing

			this.comparer = comparer;
			this.initialArraySize = initialArraySize;
			CreateInitialArray(initialArraySize);
		}

		private void CalcUsedItemsLoadFactorThreshold()
		{
			if (indexArray != null)
			{
				usedItemsLoadFactorThreshold = (int)(LoadFactor * (double)indexArray.Length);
			}
		}

		private void CreateInitialArray(int initialArraySize)
		{
			if (initialArraySize < 0)
			{
				initialArraySize = InitialItemsTableSize;
			}

			initialArray = new T[initialArraySize];
			isHashing = false;
		}

		// this code should be very similar to code when adding items into the HashTable and you know that they are unique???
		private void SwitchFromArrayToHashing()
		{
			// switch from using array into hashing
			// get the next capacity that is at least 2 * what it was before

			int newNodeArraySizeIncrease = GetNewNodeArraySizeIncrease(out int oldNodeArraySize);
			int newNodeArraySize = oldNodeArraySize + newNodeArraySizeIncrease;

			int newIndexArraySize = 0;
			if (currentIndexIntoSizeArray < indexArraySizeArray.Length - 1)
			{
				for (currentIndexIntoSizeArray = 0; currentIndexIntoSizeArray < indexArraySizeArray.Length; currentIndexIntoSizeArray++)
				{
					newIndexArraySize = indexArraySizeArray[currentIndexIntoSizeArray];
					if (newIndexArraySize >= newNodeArraySize)
					{
						break;
					}
				}
			}
			
			if (newNodeArraySize == 0)
			{
				// this is an error, the int.MaxValue has been used for capacity and we require more - throw an Exception for this???
			}

			nodeArray = new TNode[newNodeArraySize]; // the nodeArray has an extra item as it's first item (0 index) that is for available items - the memory is wasted, but it simplifies things
			indexArray = new int[newIndexArraySize]; // these will be initially set to 0, so make 0 the blank(available) value and reduce all indices by one to get to the actual index into the nodeArray
			CalcUsedItemsLoadFactorThreshold();

			// don't call AddRange here because we know for sure that these items aren't equal to each other and we can skip the equals checks
			// also we know all nodes are blank and we can avoid any logic with that as well

			int nodeIndex = 1; // start at 1 because 0 is the blank item
			for (int i = 0; i < initialArray.Length; i++)
			{
				int hash = initialArray[i].GetHashCode();
				int hashIndex = hash % newIndexArraySize;
				
				if (hashIndex < 0)
				{
					hashIndex = -hashIndex;
				}

				// idx is now from 0 to newIndexArraySize - 1

				if (indexArray[hashIndex] == 0)
				{
					nodeArray[nodeIndex].nextIndex = nodeIndex; // this node is the ending node, so set the nextIndex to itself
				}
				else
				{
					// item with same hash code already exists, so need to put in the bucket - put as first item in the bucket because it's easier
					nodeArray[nodeIndex].nextIndex = indexArray[hashIndex];
				}
				indexArray[hashIndex] = nodeIndex;
				nodeArray[nodeIndex].item = initialArray[i];
				nodeArray[nodeIndex].hashOrNextIndexForBlanks = hash;

				nodeIndex++;
			}
			nodeArray[0].nextIndex = nodeIndex; // 0 is the node used for blank - any blank item with a 0 nextIndex pointer means it really points to the next one after it
			firstBlankAtEndIndex = nodeIndex;

			initialArray = null; // this array can now be garbage collected because it is no longer referenced
			isHashing = true;
		}

		//???
		// instead of having blank items as a linked list
		// maybe indicate blank with a -1 as the nextIndex
		// and we can search for the blank ones after
		// this allows for cache locality when going through the items with the same hash code
		// so maybe create with a certain # of blank items in between each one - maybe just a single blank since we are doing 2* as the next capacity
		// could try both ways and measure

		// another way to do this would be to reformat the hashset if we know we don't need to add more items (just do lots of Contains) -   - 

		// what's the point of having 2 arrays - one for the indices and another for the Nodes/buckets???  possibly so that the index array can be resized independently of the nodes array - this allows for less collisions, but we really want nodes with the same index to be together as much as possible - this doesn't matter much for reference types, but it does for value types
		// when the index array needs to grow, we still need a good way to redo the nodes/buckets array so that items with the same index are together as much as possible in this array

		public IEqualityComparer<T> Comparer
		{
			get
			{
				return comparer;
			}
		}

		public int Count
		{
			get
			{
				return usedItemsCount;
			}
		}

		public bool IsHashing
		{
			get
			{
				return isHashing;
			}
		}

		// this is the percent of used items to all items (used + blank/available)
		// at which point (calculated ratio is >= property value) any additional added items will
		// first resize the indexArray to the next prime to avoid too many collisions and buckets becoming too large
		public double LoadFactor
		{
			get
			{
				return loadFactor;
			}

			set
			{
				loadFactor = value;
				CalcUsedItemsLoadFactorThreshold();
			}

		}

		// this is the capacity that can be trimmed with TrimExcessCapacity
		// items that were removed from the hash arrays can't be trimmed by calling TrimExcessCapacity, only the blank items at the end
		// items that were removed from the initialArray can be trimmed by calling TrimExcessCapacity because the items after are moved to fill the blank space
		public int ExcessCapacity
		{
			get
			{
				int excessCapacity;
				if (isHashing)
				{
					excessCapacity = nodeArray.Length - firstBlankAtEndIndex;
				}
				else
				{
					excessCapacity = initialArray.Length - usedItemsCount;
				}
				return excessCapacity;
			}
		}

		public int Capacity
		{
			get
			{
				if (isHashing)
				{
					return nodeArray.Length - 1; // subtract 1 for blank node at 0 index
				}
				else
				{
					return initialArray.Length;
				}
			}
		}

		// -1 means there is no set MaxCapacity, it is only limited by memory and int.MaxValue
		public int MaxCapacity { get; set; } = -1;

		// when ExcessCapacity becomes 0 and we need to allocate for more items, this overrides the next default increase, which is usually double
		// -1 indicates to use the default increase
		public int NextCapacityIncreaseOverride { get; set; } = -1;

		public int NextCapacityIncreaseDefault
		{
			get
			{
				return GetNewNodeArraySizeIncrease(out int oldNodeArraySize, true);
			}
		}

		public int NextCapacityIncrease
		{
			get
			{
				return GetNewNodeArraySizeIncrease(out int oldNodeArraySize);
			}
		}

		// allocate enough space (or make sure existing space is enough) for capacity # of items to be stored in the hashset without any further allocations
		// the actual capacity at the end of this function may be more than specified (in the case when it was more before this function was called - nothing is trimmed by this function, or in the case that slighly more capacity was allocated by this function)
		// return the actual capacity at the end of this function
        public int EnsureCapacity(int capacity)
        {
			//??? the indexArray should be allocated to be the next higher prime number using the load factor calculation (i.e. the next higher prime after the indexArray.Length / usedItemCount = LoadFactor)

			// won't this mess with that prime array and the index into it??? - not sure how to deal with this?

			return 0;
		}

		// this removes all items, but does not do any trimming of the resulting unused memory
		// to trim the unused memory, call TrimExcess
		public void Clear()
		{
			if (isHashing)
			{
				firstBlankAtEndIndex = 1;
				Array.Clear(indexArray, 0, indexArray.Length);
			}

			usedItemsCount = 0;
		}

		//
		public void ClearAndTrimAll()
		{
			// this would deallocate the arrays - would need to lazy allocate the arrays if allowing this (if (nodeArray == null) .. InitForHashing() if (initialArray == null) InitForInitialArray()
			//??? I don't think the time to always check for the arrays would make this worth it - could just set the BasicHashSet variable to null in this case and reset it with constructor when needed?
			// maybe could just check for the initialArray this way, because it would set isHashing to false?
		}

		//??? what about a function the cleans up blank internal nodes by rechaining used nodes to fill up the blank nodes
		// and then the TrimeExcess can do a better job of trimming excess - add a function to do that?  call it CompactNodeArray
		// there is a perfect structure where the first non-blank indexArray points to index 1 in the nodeArray and anything that follows does so in nodeArray index 2, 3, etc.
		// this way you have locality of reference when doing lookups and you also remove all internal blank nodes
		// it would be easy to create this structure when doing a resize of the nodeArray, so maybe doing an Array.Resize isn't the best for the nodeArray, although you are usually only doing this when you have no more blank nodes, so that part of the advantage is not valid for this scenario

		//??? I wonder how TrimExcess works for HashSet?
		public void TrimExcess()
		{
			if (isHashing)
			{
				if (nodeArray != null && nodeArray.Length > firstBlankAtEndIndex && firstBlankAtEndIndex > 0)
				{
					Array.Resize(ref nodeArray, firstBlankAtEndIndex);
					// when firstBlankAtEndIndex == nodeArray.Length, that means there are no blank at end items
				}
			}
			else
			{
				if (initialArray != null && initialArray.Length > usedItemsCount && usedItemsCount > 0)
				{
					Array.Resize(ref initialArray, usedItemsCount);
				}
			}
		}

		// return the prime number that is equal to n (if n is a prime number) or the closest prime number greather than n
		private static int GetEqualOrClosestHigherPrime(int n)
		{
			if ((n & 1) == 0)
			{
				n++; // make n odd
			}

			bool found;

			do
			{
				found = true;

				int sqrt = (int)Math.Sqrt(n);
				for (int i = 3; i < sqrt; i += 2)
				{
					if (n % i == 0)
					{
						found = false;
						n += 2;
						break;
					}
				}
			} while (!found);
 
			return n;
		}

		private int GetNewNodeArraySizeIncrease(out int oldArraySize, bool getOnlyDefaultSize = false)
		{
			if (nodeArray != null)
			{
				oldArraySize = nodeArray.Length;
			}
			else
			{
				oldArraySize = initialArray.Length; // this isn't the old node array, but it is the old # of items that could be stored without resizing
			}
				
			int increaseInSize;
			if (getOnlyDefaultSize || NextCapacityIncreaseOverride < 0)
			{
				increaseInSize = oldArraySize;
			}
			else
			{
				increaseInSize = NextCapacityIncreaseOverride;
			}

			int maxIncreaseInSize;
			if (getOnlyDefaultSize || MaxCapacity < 0)
			{
				maxIncreaseInSize = int.MaxValue - oldArraySize;
			}
			else
			{
				maxIncreaseInSize = MaxCapacity - oldArraySize;
				if (maxIncreaseInSize < 0)
				{
					maxIncreaseInSize = 0;
				}
			}

			if (increaseInSize > maxIncreaseInSize)
			{
				increaseInSize = maxIncreaseInSize;
			}
			return increaseInSize;
		}

		private int GetNewIndexArraySize()
		{
			//??? to avoid to many allocations of this array, setting the initialCapacity in the constructor or MaxCapacity or NextCapacityIncreaseOverride should determine where the currentIndexIntoSizeArray is pointing to and also the capacity of this indexArray (which should be a prime)		public int MaxCapacity { get; set; } = -1;

			int newArraySize;
			if (currentIndexIntoSizeArray < indexArraySizeArray.Length)
			{
				newArraySize = indexArraySizeArray[currentIndexIntoSizeArray];
			}
			else
			{
				newArraySize = indexArray.Length;
			}

			return newArraySize;
		}

		// if hashing, increase the size of the nodeArray
		// if not yet hashing, switch to hashing
		private void IncreaseCapacity()
		{
			if (isHashing)
			{
				int newNodeArraySizeIncrease = GetNewNodeArraySizeIncrease(out int oldNodeArraySize);

				if (newNodeArraySizeIncrease <= 0)
				{
					//??? throw an error
				}

				int newNodeArraySize = oldNodeArraySize + newNodeArraySizeIncrease;

				Array.Resize(ref nodeArray, newNodeArraySize);
				nodeArray[0].nextIndex = oldNodeArraySize;
				firstBlankAtEndIndex = oldNodeArraySize;
			}
			else
			{
				SwitchFromArrayToHashing();
			}
		}

		private void ResizeIndexArray(int newIndexArraySize)
		{
			if (newIndexArraySize == indexArray.Length)
			{
				// this will still work if no increase in size - it just might be slower than if you could increase the indexArray size
			}
			else
			{
				//??? what if there is a high percent of blank/unused items in the nodeArray before the firstBlankAtEndIndex (mabye because of lots of removes)?
				// It would probably be faster to loop through the indexArray and then do chaining to find the used nodes - one problem with this is that you would have to find blank nodes - but they would be chained
				// this probably isn't a very likely scenario
				indexArray = new int[newIndexArraySize];
				CalcUsedItemsLoadFactorThreshold();

				int indexArrayLength = indexArray.Length;

				int pastNodeIndex = nodeArray.Length;
				if (firstBlankAtEndIndex < pastNodeIndex)
				{
					pastNodeIndex = firstBlankAtEndIndex;
				}

				//??? for a loop where the end is array.Length, the compiler can skip any array bounds checking - can it do it for this code - it should be able to because pastIndex is no more than indexArray.Length
				for (int i = 1; i < pastNodeIndex; i++)
				{
					if (nodeArray[i].nextIndex != 0) // nextIndex == 0 indicates a blank/available node
					{
						int hash = nodeArray[i].hashOrNextIndexForBlanks;

						int hashIndex = hash % indexArrayLength;
						if (hashIndex < 0)
						{
							hashIndex = -hashIndex;
						}

						int nodeIndex = indexArray[hashIndex];
						if (nodeIndex == 0)
						{
							indexArray[hashIndex] = i;
							nodeArray[i].nextIndex = i;
						}
						else
						{
							indexArray[hashIndex] = i;
							nodeArray[i].nextIndex = nodeIndex;
						}
					}
				}
			}
		}

		public bool Add(T item)
		{
			if (isHashing)
			{
				return AddToHashSet(item, item.GetHashCode(), out int addedNodeIndex);
			}
			else
			{
				int i;
				for (i = 0; i < usedItemsCount; i++)
				{
					if (comparer == null ? item.Equals(initialArray[i]) : comparer.Equals(item, initialArray[i]))
					{
						return false;
					}
				}

				if (i == initialArray.Length)
				{
					SwitchFromArrayToHashing();
					return AddToHashSet(item, item.GetHashCode(), out int addedNodeIndex);
				}
				else
				{
					// add to initialArray
					initialArray[i] = item;
					usedItemsCount++;
					return true;
				}
			}
		}

		// only remove the found item if the predicate on the item evaluates to true
		//??? look at HashSet<T>.RemoveWhere
		public bool RemoveIf(T item)
		{
			bool isRemoved = false;

			return isRemoved;
		}

		//??? can we have a ref to nodeArray[0] and always use that - like blankNodeRef.nextIndex? - would this be faster? - that would probably be the only reason to do it.
		public bool Remove(T item)
		{
			bool isRemoved = false;

			if (isHashing)
			{
				FindInNodeArray(item, out int foundNodeIndex, out int priorNodeIndex, out int indexArrayIndex);
				if (foundNodeIndex > 0)
				{
					if (priorNodeIndex == 0)
					{
						if (nodeArray[foundNodeIndex].nextIndex == foundNodeIndex) // there are no more nodes in the chain - this was the only node, because the if above is only true when there is no prior node
						{
							indexArray[indexArrayIndex] = 0;
						}
						else
						{
							indexArray[indexArrayIndex] = nodeArray[foundNodeIndex].nextIndex;
							
						}
					}
					else
					{
						if (nodeArray[foundNodeIndex].nextIndex == foundNodeIndex) // the node being removed was the last node in a chain
						{
							nodeArray[priorNodeIndex].nextIndex = priorNodeIndex; // make the prior node the last node in the chain
						}
						else
						{
							nodeArray[priorNodeIndex].nextIndex = nodeArray[foundNodeIndex].nextIndex;
						}
					}

					// add node to blank chain or to the blanks at the end (if possible)
					if (foundNodeIndex == firstBlankAtEndIndex - 1)
					{
						firstBlankAtEndIndex--;
					}
					else
					{
						nodeArray[foundNodeIndex].hashOrNextIndexForBlanks = nodeArray[0].nextIndex;
						nodeArray[0].nextIndex = foundNodeIndex;
					}

					nodeArray[foundNodeIndex].nextIndex = 0;

					usedItemsCount--;
					isRemoved = true;
				}
			}
			else
			{
				for (int i = 0; i < usedItemsCount; i++)
				{
					if (comparer == null ? item.Equals(initialArray[i]) : comparer.Equals(item, initialArray[i]))
					{
						// remove the item by moving all remaining items to fill over this one
						for (int j = i + 1; j < usedItemsCount; j++, i++)
						{
							initialArray[i] = initialArray[j];
						}
						usedItemsCount--;
						isRemoved = true;
						break;
					}
				}
			}
			return isRemoved;
		}

		//??? do we need the same FindOrAdd for a reference type? or does this function take care of that?
		//??? don't do isAdded, do isFound because of Find below
		public ref T FindOrAdd(in T item, out bool isFound)
		{
			isFound = false;
			if (isHashing)
			{
				isFound = AddToHashSet(item, item.GetHashCode(), out int addedNodeIndex);
				return ref initialArray[addedNodeIndex];
			}
			else
			{
				int i;
				for (i = 0; i < usedItemsCount; i++)
				{
					if (comparer == null ? item.Equals(initialArray[i]) : comparer.Equals(item, initialArray[i]))
					{
						isFound = true;
						return ref initialArray[i];
					}
				}

				if (i == initialArray.Length)
				{
					SwitchFromArrayToHashing();
					return ref FindOrAdd(in item, out isFound);
				}
				else
				{
					// add to initialArray and keep isAdded true
					initialArray[i] = item;
					usedItemsCount++;
					return ref initialArray[i];
				}
			}
		}

		// return index into nodeArray or 0 if not found

		//??? to make things faster, could have a FindInNodeArray that just returns foundNodeIndex and another version called FindWithPriorInNodeArray that has the 3 out params
		// first test to make sure this works as is
		private void FindInNodeArray(in T item, out int foundNodeIndex, out int priorNodeIndex, out int indexArrayIndex)
		{
			foundNodeIndex = 0;
			priorNodeIndex = 0;

			int hash = item.GetHashCode();
			int hashIndex = hash % indexArray.Length;
			if (hashIndex < 0)
			{
				hashIndex = -hashIndex;
			}

			int nodeIndex = indexArray[hashIndex];
			indexArrayIndex = hashIndex;

			int priorIndex = 0;
			if (nodeIndex > 0) // 0 means item does not yet exist in the HashSet
			{
				// item with same hashIndex already exists, so need to look in the bucket for an equal item (using Equals)

				while (true)
				{
					// check if hash codes are equal before calling Equals (which may take longer) items that are Equals must have the same hash code
					if (nodeArray[nodeIndex].hashOrNextIndexForBlanks == hash && (comparer == null ? nodeArray[nodeIndex].item.Equals(item) : comparer.Equals(nodeArray[nodeIndex].item, item)))
					{
						foundNodeIndex = nodeIndex;
						priorNodeIndex = priorIndex;
						return;
					}

					int nextNodeIndex = nodeArray[nodeIndex].nextIndex;
					if (nextNodeIndex == nodeIndex)
					{
						return; // not found
					}
					else
					{
						priorIndex = nodeIndex;
						nodeIndex = nextNodeIndex;
					}
				}
			}
		}

		// this is similar to HashSet<T>.TryGetValue, except it returns a ref to the value rather than a copy of the value found (using an out parameter)
		// this way you can modify the actual value in the set if it is a value type (you can always modify the object if it is a reference type - except I think if it is a string)
		// also passing the item by in and the return by ref is faster for larger structs than passing by value
		public ref T Find(in T item, out bool isFound)
		{
			isFound = false;
			if (isHashing)
			{
				FindInNodeArray(item, out int foundNodeIndex, out int priorNodeIndex, out int indexArrayIndex);
				if (foundNodeIndex > 0)
				{
					isFound = true;
				}

				return ref nodeArray[foundNodeIndex].item;
			}
			else
			{
				int i;
				for (i = 0; i < usedItemsCount; i++)
				{
					if (comparer == null ? item.Equals(initialArray[i]) : comparer.Equals(item, initialArray[i]))
					{
						isFound = true;
						return ref initialArray[i];
					}
				}

				// if item was not found, still need to return a ref to something, so return a ref to the first item in the array
				return ref initialArray[0];
			}
		}

		public bool Contains(in T item)
		{
			bool isFound;
			Find(item, out isFound);
			return isFound;
		}

		private void IncreaseIndexArraySize()
		{
			ResizeIndexArray(GetNewIndexArraySize());
		}

		// return true if the item was added or false if it was found
		private bool AddToHashSet(T item, int hash, out int addedNodeIndex, bool checkIfItemExists = true)
		{
			int hashIndex = hash % indexArray.Length;
			if (hashIndex < 0)
			{
				hashIndex = -hashIndex;
			}

			int nodeIndex = indexArray[hashIndex];

			if (nodeIndex == 0) // 0 means item does not yet exist in the HashSet, so add it
			{
				int blankNodeIndex = nodeArray[0].nextIndex;
				if (blankNodeIndex == 0)
				{
					// 0 means there aren't any more blank nodes to add items, so we need to increase capacity
					IncreaseCapacity();
					blankNodeIndex = nodeArray[0].nextIndex;
					//return AddToHashSet(item, hash, out addedNodeIndex, false);
				}

				if (blankNodeIndex >= firstBlankAtEndIndex)
				{
					// the blank nodes starting at firstBlankAtEndIndex aren't chained
					firstBlankAtEndIndex++;
					nodeArray[0].nextIndex = firstBlankAtEndIndex;
				}
				else
				{
					// the blank nodes before firstBlankAtEndIndex are chained (the hashOrNextIndexForBlanks points to the next blank node)
					nodeArray[0].nextIndex = nodeArray[blankNodeIndex].hashOrNextIndexForBlanks;
				}
				indexArray[hashIndex] = blankNodeIndex;
				nodeArray[blankNodeIndex] = new TNode(item, blankNodeIndex, hash); // a nextIndex that equals this index indicates the end of the items with the same hashIndex, this way all blanks can be determined by a nextIndex = 0
				addedNodeIndex = blankNodeIndex;

				usedItemsCount++;

				if (usedItemsCount >= usedItemsLoadFactorThreshold)
				{
					IncreaseIndexArraySize();
				}
				return true;
			}
			else
			{
				// item with same hashIndex already exists, so need to put in the bucket if it doesn't already exist there (using Equals)

				// no need to check if item already exists if this was called recursively, since we already know it does not exist
				if (checkIfItemExists)
				{
					while (true)
					{
						// check if hash codes are equal before calling Equals (which may take longer) items that are Equals must have the same hash code
						if (nodeArray[nodeIndex].hashOrNextIndexForBlanks == hash && (comparer == null ? nodeArray[nodeIndex].item.Equals(item) : comparer.Equals(nodeArray[nodeIndex].item, item)))
						{
							addedNodeIndex = nodeIndex;
							return false;
						}

						int nextNodeIndex = nodeArray[nodeIndex].nextIndex;
						if (nextNodeIndex == nodeIndex)
						{
							break;
						}
						else
						{
							nodeIndex = nextNodeIndex;
						}
					}
				}

				int blankNodeIndex = nodeArray[0].nextIndex;
				if (blankNodeIndex == 0)
				{
					IncreaseCapacity();
					AddToHashSet(item, hash, out addedNodeIndex, false);
					return true;
				}

				nodeArray[0].nextIndex = nodeArray[blankNodeIndex].hashOrNextIndexForBlanks;
				nodeArray[blankNodeIndex] = new TNode(item, blankNodeIndex, hash); /// a nextIndex that equals this index indicates the end of the items with the same hashIndex, this way all blanks can be determined by a nextIndex = 0
				addedNodeIndex = blankNodeIndex;

				//??? add tne new node to the front, not the back, this is slightly easier, although not sure it results in faster execution 
				// it may be that more recently added items are more likely to be checked for Contains or Added in the future, so this might make that situation a little faster

				//nodeArray[nodeIndex].nextIndex = indexArray[hashIndex];
				//indexArray[hashIndex] = nodeIndex;

				usedItemsCount++;

				if (usedItemsCount >= usedItemsLoadFactorThreshold)
				{
					IncreaseIndexArraySize();
				}
				return true;
			}
		}

		public void UnionWith(IEnumerable<T> range)
		{
			//???
			// AddRange is called UnionWith in a regular HashSet - probably call it the same thing
		}

		public IEnumerator<T> GetEnumerator()
		{
			if (isHashing)
			{
				currentNodeIdx = 1;
				
				// it's easiest to just loop through the node array and skip any nodes with nextIndex = 0
				// rather than looping through the indexArray and following the nextIndex to the end of each bucket

				while (currentNodeIdx < firstBlankAtEndIndex)
				{
					if (nodeArray[currentNodeIdx].nextIndex != 0)
					{
						yield return nodeArray[currentNodeIdx].item;
					}

					currentNodeIdx++;
				}
			}
			else
			{
				currentNodeIdx = 0; // the initialArray doesn't really have nodes, but it's still just an index into the array

				while (currentNodeIdx < usedItemsCount)
				{
					yield return initialArray[currentNodeIdx];

					currentNodeIdx++;
				}
			}
		}

		IEnumerator IEnumerable.GetEnumerator()
		{
			return GetEnumerator();
		}
	}
}



