using System.Collections.Generic;
using System.Linq;


namespace Reefer_Container
{
    public static class ContainerSorter
    {
        public static IOrderedEnumerable<ReeferItem> ApplySortPriority(
            IEnumerable<ReeferItem> containers,
            IReadOnlyList<SortOrderItem> priorityItems)
        {
            // 첫 번째 우선순위 적용
            var first = priorityItems[0];
            IOrderedEnumerable<ReeferItem> ordered;

            if (first.Key == "Bay")
                ordered = first.SelectedOrder == "Fwd->Aft"
                    ? containers.OrderBy(c => c.BayNo)
                    : containers.OrderByDescending(c => c.BayNo);
            else if (first.Key == "Tier")
                ordered = first.SelectedOrder == "Top->Bottom"
                    ? containers.OrderByDescending(c => c.TierNo)
                    : containers.OrderBy(c => c.TierNo);
            else if (first.Key == "Row")
                ordered = first.SelectedOrder == "PORT->ST'BD"
                    ? containers.OrderBy(c => c.RowNo)
                    : containers.OrderByDescending(c => c.RowNo);
            else
                ordered = containers.OrderBy(c => c.BayNo);

            // 두 번째, 세 번째 우선순위 ThenBy로 체이닝
            foreach (var item in priorityItems.Skip(1))
            {
                var current = item;
                if (current.Key == "Bay")
                    ordered = current.SelectedOrder == "Fwd->Aft"
                        ? ordered.ThenBy(c => c.BayNo)
                        : ordered.ThenByDescending(c => c.BayNo);
                else if (current.Key == "Tier")
                    ordered = current.SelectedOrder == "Top->Bottom"
                        ? ordered.ThenByDescending(c => c.TierNo)
                        : ordered.ThenBy(c => c.TierNo);
                else if (current.Key == "Row")
                    ordered = current.SelectedOrder == "PORT->ST'BD"
                        ? ordered.ThenBy(c => c.RowNo)
                        : ordered.ThenByDescending(c => c.RowNo);
                else
                    ordered = ordered.ThenBy(c => c.BayNo);
            }

            return ordered;
        }
    }
}


