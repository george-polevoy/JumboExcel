using System;
using System.Collections.Generic;

namespace JumboExcel
{
    /// <summary>
    /// This collection accumulates shared elements with access by index.
    /// This class is not thread safe.
    /// </summary>
    /// <remarks>This collection is used to accumulate indexes of shared elements in OpenXml, such as <see cref="DocumentFormat.OpenXml.Spreadsheet.SharedStringTable"/></remarks>
    /// <typeparam name="T">Collection element type. This type must provide equality overrides for <see cref="IDictionary{T,Int32}"/>.</typeparam>
    class SharedElementCollection<T>
    {
        private readonly Queue<T> items = new Queue<T>();
        private readonly Dictionary<T, int> indexByItem = new Dictionary<T, int>();
        private bool done;

        /// <summary>
        /// Returns count of allocated elements.
        /// </summary>
        public int Count
        {
            get { return items.Count; }
        }

        /// <summary>
        /// Gets element index of the specified element, using the element as a dictionary key.
        /// </summary>
        /// <param name="sharedValue">Element to query.</param>
        /// <returns>Returns the allocated index.</returns>
        /// <exception cref="KeyNotFoundException">Thrown when element not found.</exception>
        public int GetElementIndex(T sharedValue)
        {
            return indexByItem[sharedValue];
        }

        /// <summary>
        /// Gets the existing index or allocates a new index, equal to the number of previously allocated elements, if the provided element is not found in collection.
        /// Uses dictionary equality to find existing element's index.
        /// If it's a new index, the instance is stored for querying and retrieval.
        /// Indexes are zero based.
        /// </summary>
        /// <param name="sharedValue">Element to query for.</param>
        /// <returns>Returns the allocated index for the provided element.</returns>
        public int GetOrAllocateElement(T sharedValue)
        {
            int index;
            if (indexByItem.TryGetValue(sharedValue, out index)) return index;
            index = items.Count;
            items.Enqueue(sharedValue);
            indexByItem.Add(sharedValue, index);
            return index;
        }

        /// <summary>
        /// Gets all existing elements, ordered by allocated indexes.
        /// </summary>
        /// <returns></returns>
        public IEnumerable<T> GetAll()
        {
            if (done)
                throw new InvalidOperationException("The collection is dequeued already.");

            return items;
        }

        /// <summary>
        /// Gets existing elements, ordered by allocated indexes. Returned value can only be enumerated once.
        /// </summary>
        /// <returns>Returns an ienumerable, which can only be enumerated once.</returns>
        /// <exception cref="InvalidOperationException">Thrown when </exception>
        public IEnumerable<T> DequeueAll()
        {
            if (done)
                throw new InvalidOperationException("The collection is dequeued already.");
            done = true;
            while (items.Count > 0)
            {
                var token = items.Dequeue();

                // it's important to remove, so that memory released to GC immidiately.
                indexByItem.Remove(token);

                yield return token;
            }
        }
    }
}
