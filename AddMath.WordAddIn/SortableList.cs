using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AddMath.WordAddIn
{
    /// <summary>
    ///  Suitable for binding to DataGridView when column sorting is required
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class SortableList<T> : BindingList<T>, IBindingListView
    {
        private PropertyComparerCollection<T> sorts;

        public SortableList()
        {
        }

        public SortableList(IEnumerable<T> initialList)
        {
            foreach (T item in initialList)
            {
                Add(item);
            }
        }

        public SortableList<T> ApplyFilter(Func<T, bool> func)
        {
            SortableList<T> newList = new SortableList<T>();
            foreach (var item in this.Where(func))
            {
                newList.Add(item);
            }

            return newList;
        }

        protected override bool IsSortedCore => sorts != null;

        protected override bool SupportsSortingCore => true;

        protected override ListSortDirection SortDirectionCore => sorts == null
                           ? ListSortDirection.Ascending
                           : sorts.PrimaryDirection;

        protected override PropertyDescriptor SortPropertyCore => sorts?.PrimaryProperty;

        public void ApplySort(ListSortDescriptionCollection sortCollection)
        {
            bool oldRaise = RaiseListChangedEvents;
            RaiseListChangedEvents = false;
            try
            {
                PropertyComparerCollection<T> tmp
                    = new PropertyComparerCollection<T>(sortCollection);
                List<T> items = new List<T>(this);
                items.Sort(tmp);
                int index = 0;
                foreach (T item in items)
                {
                    SetItem(index++, item);
                }
                sorts = tmp;
            }
            finally
            {
                RaiseListChangedEvents = oldRaise;
                ResetBindings();
            }
        }

        string IBindingListView.Filter
        {
            get { throw new NotImplementedException(); }
            set { throw new NotImplementedException(); }
        }

        void IBindingListView.RemoveFilter()
        {
            throw new NotImplementedException();
        }

        ListSortDescriptionCollection IBindingListView.SortDescriptions => sorts?.Sorts;

        bool IBindingListView.SupportsAdvancedSorting => true;

        bool IBindingListView.SupportsFiltering => false;

        protected override void RemoveSortCore() => sorts = null;

        protected override void ApplySortCore(PropertyDescriptor prop, ListSortDirection direction)
        {
            ListSortDescription[] arr = { new ListSortDescription(prop, direction) };
            ApplySort(new ListSortDescriptionCollection(arr));
        }
    }

    public class PropertyComparerCollection<T> : IComparer<T>
    {
        private readonly PropertyComparer<T>[] comparers;

        private readonly ListSortDescriptionCollection sorts;

        public PropertyComparerCollection(ListSortDescriptionCollection
                                              sorts)
        {
            if (sorts == null)
            {
                throw new ArgumentNullException("sorts");
            }
            this.sorts = sorts;
            List<PropertyComparer<T>> list = new
                List<PropertyComparer<T>>();

            foreach (ListSortDescription item in sorts)
            {
                list.Add(new PropertyComparer<T>(item.PropertyDescriptor,
                                                 item.SortDirection == ListSortDirection.Descending));
            }

            comparers = list.ToArray();
        }

        public ListSortDescriptionCollection Sorts => sorts;

        public PropertyDescriptor PrimaryProperty => comparers.Length == 0
                           ? null
                           : comparers[0].Property;

        public ListSortDirection PrimaryDirection => comparers.Length == 0
                           ? ListSortDirection.Ascending
                           : comparers[0].Descending
                                 ? ListSortDirection.Descending
                                 : ListSortDirection.Ascending;

        int IComparer<T>.Compare(T x, T y)
        {
            int result = 0;
            foreach (PropertyComparer<T> t in comparers)
            {
                result = t.Compare(x, y);
                if (result != 0)
                {
                    break;
                }
            }
            return result;
        }
    }

    public class PropertyComparer<T> : IComparer<T>
    {
        private readonly bool descending;

        private readonly PropertyDescriptor property;

        public PropertyComparer(PropertyDescriptor property, bool descending)
        {
            if (property == null)
            {
                throw new ArgumentNullException("property");
            }

            this.descending = descending;
            this.property = property;
        }

        public bool Descending => descending;

        public PropertyDescriptor Property => property;

        public int Compare(T x, T y)
        {
            int value = Comparer.Default.Compare(property.GetValue(x),
                                                 property.GetValue(y));
            return descending ? -value : value;
        }
    }   
}
