// ---------------------------------------------------------------------------------------------------------------
// Name:     TabList.cs
// Purpose:  The ordered list of tabs for one app, plus which tab is current. Used on the Office UI thread only.
//           Every change raises Changed once, after the same-name suffixes are recalculated.
// Revision: 2
// DTG:      261024200000
// SPDX-License-Identifier: MIT
// ---------------------------------------------------------------------------------------------------------------

using System;
using System.Collections.Generic;

namespace OpenTools
{
    internal sealed class TabList
    {
        private readonly List<TabItem> _items = new List<TabItem>();
        private string _currentKey = string.Empty;

        public event EventHandler Changed;

        // ----- Reading -----

        public IList<TabItem> Items
        {
            get { return _items.AsReadOnly(); }
        }

        public int Count
        {
            get { return _items.Count; }
        }

        public string CurrentKey
        {
            get { return _currentKey; }
        }

        public TabItem Find(string key)
        {
            foreach (TabItem item in _items)
            {
                if (item.Key == key)
                {
                    return item;
                }
            }

            return null;
        }

        public int IndexOf(string key)
        {
            for (int i = 0; i < _items.Count; i++)
            {
                if (_items[i].Key == key)
                {
                    return i;
                }
            }

            return -1;
        }

        // ----- Changing -----

        public void Add(TabItem item)
        {
            if (Find(item.Key) != null)
            {
                return;
            }

            _items.Add(item);
            RaiseChanged();
        }

        public void Remove(string key)
        {
            int index = IndexOf(key);
            if (index < 0)
            {
                return;
            }

            _items.RemoveAt(index);
            if (_currentKey == key)
            {
                _currentKey = string.Empty;
            }

            RaiseChanged();
        }

        public void SetCurrent(string key)
        {
            if (_currentKey == key)
            {
                return;
            }

            _currentKey = key;
            RaiseChanged();
        }

        public void Move(int fromIndex, int toIndex)
        {
            if (fromIndex < 0 || fromIndex >= _items.Count || toIndex < 0 || toIndex >= _items.Count
                || fromIndex == toIndex)
            {
                return;
            }

            TabItem item = _items[fromIndex];
            _items.RemoveAt(fromIndex);
            _items.Insert(toIndex, item);
            RaiseChanged();
        }

        public void UpdateFile(string key, string fullPath, string name)
        {
            TabItem item = Find(key);
            if (item == null || (item.FullPath == fullPath && item.Name == name))
            {
                return;
            }

            item.FullPath = fullPath;
            item.Name = name;
            RaiseChanged();
        }

        public void SetDirty(string key, bool isDirty)
        {
            TabItem item = Find(key);
            if (item == null || item.IsDirty == isDirty)
            {
                return;
            }

            item.IsDirty = isDirty;
            RaiseChanged();
        }

        private void RaiseChanged()
        {
            TabLabeler.AssignSuffixes(_items);
            EventHandler handler = Changed;
            if (handler != null)
            {
                handler(this, EventArgs.Empty);
            }
        }
    }
}
