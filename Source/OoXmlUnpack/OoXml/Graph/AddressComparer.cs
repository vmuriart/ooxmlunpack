namespace OoXml.Graph
{
    using System;
    using System.Collections.Generic;
    using System.Linq;

    using Com.AriadneInsight.Schematiq.Parser.Addresses;

    internal class AddressComparer : IComparer<IAddress>
    {
        public int Compare(IAddress x, IAddress y)
        {
            var xAddresses = x.Addresses.ToList();
            var yAddresses = y.Addresses.ToList();

            var minLength = xAddresses.Count < yAddresses.Count ? xAddresses.Count : yAddresses.Count;
            for (var i = 0; i < minLength; i++)
            {
                var xAddress = xAddresses[i];
                var yAddress = yAddresses[i];

                var compareWorkbook = string.Compare(
                    xAddress.Workbook,
                    yAddress.Workbook,
                    StringComparison.InvariantCultureIgnoreCase);
                if (compareWorkbook != 0)
                {
                    return compareWorkbook;
                }

                var compareWorksheet = string.Compare(xAddress.Worksheet, yAddress.Worksheet, StringComparison.OrdinalIgnoreCase);
                if (compareWorksheet != 0)
                {
                    return compareWorksheet;
                }

                var compareRow = xAddress.Row.CompareTo(yAddress.Row);
                if (compareRow != 0)
                {
                    return compareRow;
                }

                var compareColumn = xAddress.Column.CompareTo(yAddress.Column);
                if (compareColumn != 0)
                {
                    return compareColumn;
                }
            }

            return xAddresses.Count.CompareTo(yAddresses.Count);
        }
    }
}