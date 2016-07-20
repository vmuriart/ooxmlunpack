namespace OoXmlUnpack.Graph
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Xml.Linq;

    using Com.AriadneInsight.Schematiq.Parser;
    using Com.AriadneInsight.Schematiq.Parser.Addresses;
    using Com.AriadneInsight.Schematiq.Parser.Parsing;
    using Com.AriadneInsight.Schematiq.Parser.Parsing.Nodes;

    internal class SheetToGraph
    {
        private readonly FileInfo file;

        public SheetToGraph(FileInfo file)
        {
            this.file = file;
        }

        public void ConvertToCode()
        {
            var doc = XDocument.Load(this.file.FullName);

            var cellFormulas = new SortedDictionary<IAddress, ExcelFormula<A1Style>>(new AddressComparer());
            var cellElements = new Dictionary<IAddress, XElement>();
            var ns = doc.Root.Name.Namespace;
            if (doc.Root.Name.LocalName != "worksheet")
            {
                return;
            }

            var sheetData = doc.Descendants(ns + "sheetData").SingleOrDefault();
            if (sheetData == null)
            {
                return;
            }

            var toRemove = new List<XElement>();
            foreach (var row in sheetData.Descendants(ns + "row"))
            {
                foreach (var cell in row.Elements(ns + "c"))
                {
                    var r = cell.Attribute("r");
                    var address = Formula.ParseA1(r.Value).ToAddress(Address.Worksheet("Book1", "Sheet1"));
                    cellElements.Add(address, cell);

                    var f = cell.Element(ns + "f");
                    if (f != null && !string.IsNullOrWhiteSpace(f.Value))
                    {
                        try
                        {
                            cellFormulas.Add(address, Formula.ParseA1(f.Value));
                        }
                        catch
                        {
                            cellFormulas.Add(address, null);
                        }
                    }
                    else
                    {
                        cellFormulas.Add(address, null);
                    }
                }

                toRemove.Add(row);
            }

            toRemove.ForEach(x => x.Remove());

            var nodes = new Dictionary<IAddress, GraphNode<IAddress, ExcelFormula<A1Style>>>();
            while (cellFormulas.Count > 0)
            {
                var current = cellFormulas.First();

                // We want to find columns of cells which contain the same (R1C1) formula
                var rows = 0;
                IAddress cellBelow;
                ExcelFormula<A1Style> formulaBelow;
                do
                {
                    rows++;
                    cellBelow = current.Key[rows, 0];
                }
                while (current.Value != null && cellFormulas.TryGetValue(cellBelow, out formulaBelow)
                       && formulaBelow != null
                       && formulaBelow.ToR1C1Style(cellBelow).Equals(current.Value.ToR1C1Style(current.Key)));

                var bottomRight = current.Key.Resize(rows, 1).BottomRight;
                for (var i = 0; i < rows; i++)
                {
                    var cellAddress = current.Key[i, 0];
                    var multiAddress = bottomRight.Union(cellAddress);
                    nodes.Add(cellAddress, new GraphNode<IAddress, ExcelFormula<A1Style>>(multiAddress, cellFormulas[cellAddress]));
                    cellFormulas.Remove(cellAddress);
                }
            }

            var noPrecedents = new SortedList<IAddress, GraphNode<IAddress, ExcelFormula<A1Style>>>(new AddressComparer());
            foreach (var node in nodes)
            {
                if (node.Value.Value == null)
                {
                    noPrecedents.Add(node.Value.Key, node.Value);
                    continue;
                }

                var anyPrecedents = false;
                var parsed = node.Value.Value;
                foreach (
                    var referencedAddress in
                        GetReferences(parsed, node.Key).Where(x => !x.IsName).SelectMany(x => x.Addresses))
                {
                    for (var i = 0; i < referencedAddress.Rows; i++)
                    {
                        for (var j = 0; j < referencedAddress.Columns; j++)
                        {
                            var referencedCell = referencedAddress[i, j];
                            GraphNode<IAddress, ExcelFormula<A1Style>> referencedFormula;
                            if (nodes.TryGetValue(referencedCell, out referencedFormula))
                            {
                                node.Value.DependsOn(referencedFormula);
                                anyPrecedents = true;
                            }
                        }
                    }
                }

                if (!anyPrecedents)
                {
                    noPrecedents.Add(node.Value.Key, node.Value);
                }
            }

            var sorted = new List<GraphNode<IAddress, ExcelFormula<A1Style>>>();
            while (noPrecedents.Count > 0)
            {
                var next = noPrecedents.Values[0];

                noPrecedents.Remove(next.Key);

                sorted.Add(next);
                foreach (var dependent in next.Dependents.Values)
                {
                    dependent.Precedents.Remove(next.Key);
                    if (dependent.Precedents.Count == 0)
                    {
                        noPrecedents.Add(dependent.Key, dependent);
                    }
                }
            }

            var styleMap = new Dictionary<string, List<IAddress>>();
            var styleOrder = new List<XAttribute>();
            var cellsElement = new XElement(ns + "cells");
            var previousAddress = Address.Range("Book1", "Sheet1", new OneBasedIndex(1), new OneBasedIndex(1));
            foreach (var node in sorted)
            {
                var cellAddress = node.Key.Addresses.Last();
                var cellElement = cellElements[cellAddress];
                var styleAttribute = cellElement.Attribute("s");
                if (styleAttribute != null)
                {
                    styleAttribute.Remove();
                    List<IAddress> styleCells;
                    if (styleMap.TryGetValue(styleAttribute.Value, out styleCells))
                    {
                        styleCells.Add(node.Key);
                    }
                    else
                    {
                        styleOrder.Add(styleAttribute);
                        styleMap.Add(styleAttribute.Value, new List<IAddress> { node.Key });
                    }
                }

                if (cellElement.HasElements || cellElement.Attributes().Any(x => x.Name != "r"))
                {
                    var r = cellElement.Attribute("r");
                    cellElement.Add(new XAttribute("r2", cellAddress.ReferencedFromR1C1(previousAddress).ToString()));
                    r.Remove();
                    previousAddress = cellAddress;

                    if (node.Value != null)
                    {
                        var formulaElement = cellElement.Element(ns + "f");
                        if (formulaElement != null)
                        {
                            formulaElement.Remove();
                        }

                        cellElement.Add(new XElement(ns + "f2", node.Value.ToR1C1Style(cellAddress).ToString()));
                    }

                    cellsElement.Add(cellElement);
                }
            }

            var cellsFile = Path.Combine(this.file.Directory.FullName, string.Format("{0}-cells.xml", Path.GetFileNameWithoutExtension(this.file.Name)));

            new XDocument(cellsElement).Save(cellsFile);

            var stylesElement = new XElement(ns + "styles");
            foreach (var style in styleOrder)
            {
                var styleElement = new XElement(ns + "style", style);
                stylesElement.Add(styleElement);
                foreach (var cell in styleMap[style.Value])
                {
                    styleElement.Add(new XElement(ns + "c", new XAttribute("r", cell.Addresses.Last().ReferencedFrom(cell.EntireSheet).ToString())));
                }
            }

            var stylesFile = Path.Combine(this.file.Directory.FullName, string.Format("{0}-styles.xml", Path.GetFileNameWithoutExtension(this.file.Name)));

            new XDocument(stylesElement).Save(stylesFile);

            doc.Save(Path.Combine(this.file.Directory.FullName, string.Format("{0}-core.xml", Path.GetFileNameWithoutExtension(this.file.Name))));
        }

        private static IEnumerable<IAddress> GetReferences(ExcelFormula<A1Style> formula, IAddress referencedFrom)
        {
            switch (formula.NodeType & FormulaNodeType.CategoryMask)
            {
                case FormulaNodeType.BinaryOperator:
                    switch (formula.NodeType)
                    {
                        case FormulaNodeType.RangeOperator:
                        case FormulaNodeType.UnionOperator:
                        case FormulaNodeType.IntersectOperator:
                            return new[] { formula.ToAddress(referencedFrom) };
                        default:
                            return
                                new[] { formula.BinaryOperator.FirstArgument, formula.BinaryOperator.SecondArgument }.SelectMany
                                    (x => GetReferences(x, referencedFrom));
                    }
                case FormulaNodeType.UnaryOperator:
                    return GetReferences(formula.UnaryOperator.Argument, referencedFrom);
                case FormulaNodeType.FunctionCall:
                    return formula.Function.Arguments.SelectMany(x => GetReferences(x, referencedFrom));
                case FormulaNodeType.Literal:
                    return new IAddress[0];
                case FormulaNodeType.Reference:
                    return new[] { formula.ToAddress(referencedFrom) };
                default:
                    throw new ApplicationException("Unexpected node type");
            }
        }
    }
}