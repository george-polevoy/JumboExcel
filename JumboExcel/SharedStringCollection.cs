using System;
using System.Collections.Generic;

namespace JumboExcel
{
    class SharedElementCollection<T>
    {
        private readonly Queue<T> items = new Queue<T>();
        private readonly Dictionary<T, int> indexByItem = new Dictionary<T, int>();

        private bool done;

        public int Count
        {
            get { return items.Count; }
        }

        public int GetElementIndex(T sharedValue)
        {
            return indexByItem[sharedValue];
        }

        public int AllocateElement(T sharedValue)
        {
            int sharedStringIndex;
            if (indexByItem.TryGetValue(sharedValue, out sharedStringIndex)) return sharedStringIndex;
            sharedStringIndex = items.Count;
            items.Enqueue(sharedValue);
            indexByItem.Add(sharedValue, sharedStringIndex);
            return sharedStringIndex;
        }

        public IEnumerable<T> GetAll()
        {
            return items;
        }

        public IEnumerable<T> DequeueAll()
        {
            if (done)
                throw new InvalidOperationException("It's done already.");
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
