// .NET COM interop strongly typed collection generic implementation classes
// March 2014
// Paul Voelker

using System;
using System.Collections;

namespace COMHelpers
{
    /// <summary>
    /// Generic read-only wrapper class for COM collections
    /// </summary>
    /// <typeparam name="T">COM implementation class</typeparam>
    public class GenericReadOnlyListImplementationCOM<T>
    {
        ArrayList m_coll = null;

        public GenericReadOnlyListImplementationCOM(ArrayList coll) { m_coll = coll; }
        public GenericReadOnlyListImplementationCOM() { m_coll = new ArrayList(); }

        public static implicit operator ArrayList(GenericReadOnlyListImplementationCOM<T> value) { return value.m_coll; } // Allows for implicit downcasting to ArrayList

        public int Count { get { return m_coll.Count; } }

        public T GetItem(int index) { return (T)m_coll[index]; }
        public int IndexOf(T item) { return m_coll.IndexOf(item); }
        public bool Contains(T item) { return m_coll.Contains(item); }
    }

    /// <summary>
    /// Generic wrapper class for COM collections
    /// </summary>
    /// <typeparam name="T">COM implementation class</typeparam>
    public class GenericListImplementationCOM<T>
    {
        ArrayList m_coll = null;

        public GenericListImplementationCOM(ArrayList coll) { m_coll = coll; }
        public GenericListImplementationCOM() { m_coll = new ArrayList(); }

        public static implicit operator ArrayList(GenericListImplementationCOM<T> value) { return value.m_coll; } // Allows for implicit downcasting to ArrayList

        public int Count { get { return m_coll.Count; } }

        public void Clear() { m_coll.Clear(); }

        public void Add(T item) { m_coll.Add(item); }
        public void Insert(int index, T item) { m_coll.Insert(index, item); }

        public void RemoveAt(int index) { m_coll.RemoveAt(index); }
        public void Remove(T item) { m_coll.Remove((object)item); }

        public T GetItem(int index) { return (T)m_coll[index]; }
        public int IndexOf(T item) { return m_coll.IndexOf(item); }
        public bool Contains(T item) { return m_coll.Contains(item); }
    }
}
