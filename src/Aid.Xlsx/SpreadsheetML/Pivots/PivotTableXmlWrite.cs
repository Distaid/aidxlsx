using System.Diagnostics;
using System.Xml;
using Array = System.Array;
using static Aid.Xlsx.SpreadsheetML.Xml.OpenXmlConst;
using static Aid.Xlsx.Model.Workbooks.XlsxWorkbook;
using Aid.Xlsx.Infrastructure.Extensions;
using Aid.Xlsx.Infrastructure.Utilities;
using Aid.Xlsx.Model.PivotTables;
using Aid.Xlsx.Model.PivotTables.Areas;
using Aid.Xlsx.Model.PivotTables.PivotValues;
using Aid.Xlsx.Model.Workbooks;

namespace Aid.Xlsx.SpreadsheetML.Pivots;

internal static partial class PivotTableXml
{
    internal static void Write(Stream stream, XlsxPivotTable pt, SaveContext context)
    {
        var settings = new XmlWriterSettings
        {
            Encoding = XlsxHelper.NoBomUTF8
        };

        using var xml = XmlWriter.Create(stream, settings);

        xml.WriteStartDocument();
        xml.WriteStartElement("pivotTableDefinition", Main2006SsNs);
        xml.WriteAttributeString("xmlns", Main2006SsNs);

        xml.WriteAttribute("name", pt.Name);
        xml.WriteAttribute("cacheId", pt.PivotCache.CacheId!.Value); // TODO: Maybe not nullable?
        xml.WriteAttributeDefault("dataOnRows", pt.DataOnRows, false);
        xml.WriteAttributeOptional("dataPosition", pt.DataPosition);
        xml.WriteAttributeOptional("autoFormatId", pt.AutoFormatId);

        // Although apply*Formats do have default value `false`, Excel always writes them.
        xml.WriteAttribute("applyNumberFormats", pt.ApplyNumberFormats);
        xml.WriteAttribute("applyBorderFormats", pt.ApplyBorderFormats);
        xml.WriteAttribute("applyFontFormats", pt.ApplyFontFormats);
        xml.WriteAttribute("applyPatternFormats", pt.ApplyPatternFormats);
        xml.WriteAttribute("applyAlignmentFormats", pt.ApplyAlignmentFormats);
        xml.WriteAttribute("applyWidthHeightFormats", pt.ApplyWidthHeightFormats);

        xml.WriteAttribute("dataCaption", pt.DataCaption);
        xml.WriteAttributeOptional("grandTotalCaption", pt.GrandTotalCaption);
        xml.WriteAttributeOptional("errorCaption", pt.ErrorValueReplacement);
        xml.WriteAttributeDefault("showError", pt.ShowError, false);
        xml.WriteAttributeOptional("missingCaption", pt.MissingCaption);
        xml.WriteAttributeDefault("showMissing", pt.ShowMissing, true);
        xml.WriteAttributeOptional("pageStyle", pt.PageStyle);
        xml.WriteAttributeOptional("pivotTableStyle", pt.PivotTableStyleName);
        xml.WriteAttributeOptional("vacatedStyle", pt.VacatedStyle);
        xml.WriteAttributeOptional("tag", pt.Tag);
        xml.WriteAttributeDefault("updatedVersion", pt.UpdatedVersion, 0);
        xml.WriteAttributeDefault("minRefreshableVersion", pt.MinRefreshableVersion, 0);
        xml.WriteAttributeDefault("asteriskTotals", pt.AsteriskTotals, false);
        xml.WriteAttributeDefault("showItems", pt.DisplayItemLabels, true);
        xml.WriteAttributeDefault("editData", pt.EditData, false);
        xml.WriteAttributeDefault("disableFieldList", pt.DisableFieldList, false);
        xml.WriteAttributeDefault(@"showCalcMbrs", pt.ShowCalculatedMembers, true);
        xml.WriteAttributeDefault("visualTotals", pt.VisualTotals, true);
        xml.WriteAttributeDefault("showMultipleLabel", pt.ShowMultipleLabel, true);
        xml.WriteAttributeDefault("showDataDropDown", pt.ShowDataDropDown, true);
        xml.WriteAttributeDefault("showDrill", pt.ShowExpandCollapseButtons, true);
        xml.WriteAttributeDefault("printDrill", pt.PrintExpandCollapsedButtons, false);
        xml.WriteAttributeDefault("showMemberPropertyTips", pt.ShowPropertiesInTooltips, true);
        xml.WriteAttributeDefault("showDataTips", pt.ShowContextualTooltips, true);
        xml.WriteAttributeDefault("enableWizard", pt.EnableEditingMechanism, true);
        xml.WriteAttributeDefault("enableDrill", pt.EnableShowDetails, true);
        xml.WriteAttributeDefault("enableFieldProperties", pt.EnableFieldProperties, true);
        xml.WriteAttributeDefault("preserveFormatting", pt.PreserveCellFormatting, true);
        xml.WriteAttributeDefault("useAutoFormatting", pt.AutofitColumns, false);
        xml.WriteAttributeDefault("pageWrap", checked((uint)pt.FilterFieldsPageWrap), 0);
        xml.WriteAttributeDefault("pageOverThenDown", pt.FilterAreaOrder == XlsxFilterAreaOrder.OverThenDown, false);
        xml.WriteAttributeDefault("subtotalHiddenItems", pt.FilteredItemsInSubtotals, false);
        xml.WriteAttributeDefault("rowGrandTotals", pt.ShowGrandTotalsRows, true);
        xml.WriteAttributeDefault("colGrandTotals", pt.ShowGrandTotalsColumns, true);
        xml.WriteAttributeDefault("fieldPrintTitles", pt.PrintTitles, false);
        xml.WriteAttributeDefault("itemPrintTitles", pt.RepeatRowLabels, false);
        xml.WriteAttributeDefault("mergeItem", pt.MergeAndCenterWithLabels, false);
        xml.WriteAttributeDefault("showDropZones", pt.ShowDropZones, true);
        xml.WriteAttributeDefault("createdVersion", pt.PivotCacheCreatedVersion, 0);
        xml.WriteAttributeDefault("indent", checked((uint)pt.RowLabelIndent), 1);
        xml.WriteAttributeDefault("showEmptyRow", pt.ShowEmptyItemsOnRows, false);
        xml.WriteAttributeDefault("showEmptyCol", pt.ShowEmptyItemsOnColumns, false);
        xml.WriteAttributeDefault("showHeaders", pt.DisplayCaptionsAndDropdowns, true);
        xml.WriteAttributeDefault("compact", pt.Compact, true);
        xml.WriteAttributeDefault("outline", pt.Outline, false);
        xml.WriteAttributeDefault("outlineData", pt.OutlineData, false);
        xml.WriteAttributeDefault("compactData", pt.CompactData, true);
        xml.WriteAttributeDefault("published", pt.Published, false);
        xml.WriteAttributeDefault("gridDropZones", pt.ClassicPivotTableLayout, false);
        xml.WriteAttributeDefault("immersive", pt.StopImmersiveUi, true);
        xml.WriteAttributeDefault("multipleFieldFilters", pt.AllowMultipleFilters, true);
        xml.WriteAttributeDefault("chartFormat", pt.ChartFormat, 0);
        xml.WriteAttributeOptional("rowHeaderCaption", pt.RowHeaderCaption);
        xml.WriteAttributeOptional("colHeaderCaption", pt.ColumnHeaderCaption);
        xml.WriteAttributeDefault("fieldListSortAscending", pt.SortFieldsAtoZ, false);
        xml.WriteAttributeDefault(@"mdxSubqueries", pt.MdxSubQueries, false);
        xml.WriteAttributeDefault("customListSort", pt.UseCustomListsForSorting, true);

        // Location
        xml.WriteStartElement("location", Main2006SsNs);
        xml.WriteAttribute("ref", pt.Area.ToString());
        xml.WriteAttribute("firstHeaderRow", pt.FirstHeaderRow);
        xml.WriteAttribute("firstDataRow", pt.FirstDataRow);
        xml.WriteAttribute("firstDataCol", pt.FirstDataCol);

        (int Width, int Height) = pt.Filters.GetSize();
        xml.WriteAttributeDefault("rowPageCount", Height, 0);
        xml.WriteAttributeDefault("colPageCount", Width, 0);
        xml.WriteEndElement(); // location

        // Pivot Fields
        xml.WriteStartElement("pivotFields", Main2006SsNs);
        xml.WriteAttribute("count", pt.PivotFields.Count);

        foreach (XlsxPivotTableField? pf in pt.PivotFields)
        {
            xml.WriteStartElement("pivotField", Main2006SsNs);
            xml.WriteAttributeOptional("name", pf.Name);

            if (pf.Axis is not null)
            {
                var axisAttr = GetAxisAttr(pf.Axis.Value);
                xml.WriteAttribute("axis", axisAttr);
            }

            xml.WriteAttributeDefault("dataField", pf.DataField, false);
            xml.WriteAttributeOptional("subtotalCaption", pf.SubtotalCaption);
            xml.WriteAttributeDefault("showDropDowns", pf.ShowDropDowns, true);
            xml.WriteAttributeDefault("hiddenLevel", pf.HiddenLevel, false);
            xml.WriteAttributeOptional("uniqueMemberProperty", pf.UniqueMemberProperty);
            xml.WriteAttributeDefault("compact", pf.Compact, true);
            xml.WriteAttributeDefault("allDrilled", pf.AllDrilled, false);
            xml.WriteAttributeOptional("numFmtId", context.GetNumberFormatId(pf.NumberFormatValue));
            xml.WriteAttributeDefault("outline", pf.Outline, true);
            xml.WriteAttributeDefault("subtotalTop", pf.SubtotalTop, true);
            xml.WriteAttributeDefault("dragToRow", pf.DragToRow, true);
            xml.WriteAttributeDefault("dragToCol", pf.DragToColumn, true);
            xml.WriteAttributeDefault("multipleItemSelectionAllowed", pf.MultipleItemSelectionAllowed, false);
            xml.WriteAttributeDefault("dragToPage", pf.DragToPage, true);
            xml.WriteAttributeDefault("dragToData", pf.DragToData, true);
            xml.WriteAttributeDefault("dragOff", pf.DragOff, true);
            xml.WriteAttributeDefault("showAll", pf.ShowAll, true);
            xml.WriteAttributeDefault("insertBlankRow", pf.InsertBlankRow, false);
            xml.WriteAttributeDefault("serverField", pf.ServerField, false);
            xml.WriteAttributeDefault("insertPageBreak", pf.InsertPageBreak, false);
            xml.WriteAttributeDefault("autoShow", pf.AutoShow, false);
            xml.WriteAttributeDefault("topAutoShow", pf.TopAutoShow, true);
            xml.WriteAttributeDefault("hideNewItems", pf.HideNewItems, false);
            xml.WriteAttributeDefault("measureFilter", pf.MeasureFilter, false);
            xml.WriteAttributeDefault("includeNewItemsInFilter", pf.IncludeNewItemsInFilter, false);
            xml.WriteAttributeDefault("itemPageCount", pf.ItemPageCount, 10);
            if (pf.SortType != XlsxPivotSortType.Default)
            {
                var sortTypeAttr = pf.SortType switch
                {
                    XlsxPivotSortType.Default => "manual",
                    XlsxPivotSortType.Ascending => "ascending",
                    XlsxPivotSortType.Descending => "descending",
                    _ => throw new UnreachableException(),
                };
                xml.WriteAttribute("sortType", sortTypeAttr);
            }

            xml.WriteAttributeOptional("dataSourceSort", pf.DataSourceSort);
            xml.WriteAttributeDefault("nonAutoSortDefault", pf.NonAutoSortDefault, false);
            xml.WriteAttributeOptional("rankBy", pf.RankBy);
            xml.WriteAttributeDefault("defaultSubtotal", pf.Subtotals.Contains(XlsxSubtotalFunction.Automatic), true);
            xml.WriteAttributeDefault("sumSubtotal", pf.Subtotals.Contains(XlsxSubtotalFunction.Sum), false);
            xml.WriteAttributeDefault("countASubtotal", pf.Subtotals.Contains(XlsxSubtotalFunction.Count), false);
            xml.WriteAttributeDefault("avgSubtotal", pf.Subtotals.Contains(XlsxSubtotalFunction.Average), false);
            xml.WriteAttributeDefault("maxSubtotal", pf.Subtotals.Contains(XlsxSubtotalFunction.Maximum), false);
            xml.WriteAttributeDefault("minSubtotal", pf.Subtotals.Contains(XlsxSubtotalFunction.Minimum), false);
            xml.WriteAttributeDefault("productSubtotal", pf.Subtotals.Contains(XlsxSubtotalFunction.Product), false);
            xml.WriteAttributeDefault("countSubtotal", pf.Subtotals.Contains(XlsxSubtotalFunction.CountNumbers), false);
            xml.WriteAttributeDefault("stdDevSubtotal", pf.Subtotals.Contains(XlsxSubtotalFunction.StandardDeviation), false);
            xml.WriteAttributeDefault("stdDevPSubtotal", pf.Subtotals.Contains(XlsxSubtotalFunction.PopulationStandardDeviation), false);
            xml.WriteAttributeDefault("varSubtotal", pf.Subtotals.Contains(XlsxSubtotalFunction.Variance), false);
            xml.WriteAttributeDefault("varPSubtotal", pf.Subtotals.Contains(XlsxSubtotalFunction.PopulationVariance), false);
            xml.WriteAttributeDefault("showPropCell", pf.ShowPropCell, false);
            xml.WriteAttributeDefault("showPropTip", pf.ShowPropTip, false);
            xml.WriteAttributeDefault("showPropAsCaption", pf.ShowPropAsCaption, false);
            xml.WriteAttributeDefault("defaultAttributeDrillState", pf.DefaultAttributeDrillState, false);

            // items
            if (pf.Items.Count > 0)
            {
                xml.WriteStartElement("items", Main2006SsNs);
                xml.WriteAttribute("count", pf.Items.Count);
                foreach (XlsxPivotFieldItem pfItem in pf.Items)
                {
                    xml.WriteStartElement("item", Main2006SsNs);
                    xml.WriteAttributeOptional("n", pfItem.ItemUserCaption);
                    if (pfItem.ItemType != XlsxPivotItemType.Data)
                    {
                        var itemTypeAttr = GetItemTypeAttr(pfItem.ItemType);
                        xml.WriteAttribute("t", itemTypeAttr);
                    }

                    xml.WriteAttributeDefault("h", pfItem.Hidden, false);
                    xml.WriteAttributeDefault("s", pfItem.ValueIsString, false);
                    xml.WriteAttributeDefault("sd", pfItem.ShowDetails, true);
                    xml.WriteAttributeDefault("f", pfItem.CalculatedMember, false);
                    xml.WriteAttributeDefault("m", pfItem.Missing, false);
                    xml.WriteAttributeDefault("c", pfItem.ApproximatelyHasChildren, false);
                    xml.WriteAttributeOptional("x", pfItem.ItemIndex);
                    xml.WriteAttributeDefault("d", pfItem.Details, false);
                    xml.WriteAttributeDefault("e", pfItem.DrillAcrossAttributes, true);
                    xml.WriteEndElement(); // item
                }

                xml.WriteEndElement(); // items
            }

            // TODO: autoSortScope, but not yet represented.

            if (pf.RepeatItemLabels)
            {
                xml.WriteStartElement("extLst");
                xml.WriteStartElement("ext");
                xml.WriteAttributeString("uri", "{2946ED86-A175-432a-8AC1-64E0C546D7DE}");
                xml.WriteStartElement("pivotField", X14Main2009SsNs);
                xml.WriteAttributeDefault("fillDownLabels", pf.RepeatItemLabels, false);
                xml.WriteEndElement(); // pivotField
                xml.WriteEndElement(); // ext
                xml.WriteEndElement(); // extLst
            }

            xml.WriteEndElement();
        }

        xml.WriteEndElement(); // pivotFields

        WriteAxis(xml, pt.RowAxis, "rowFields", "rowItems");
        WriteAxis(xml, pt.ColumnAxis, "colFields", "colItems");

        IReadOnlyList<XlsxPivotPageField> filterFields = pt.Filters.Fields;
        if (filterFields.Count > 0)
        {
            xml.WriteStartElement("pageFields", Main2006SsNs);
            xml.WriteAttribute("count", filterFields.Count);
            foreach (XlsxPivotPageField filterField in filterFields)
            {
                xml.WriteStartElement("pageField", Main2006SsNs);
                xml.WriteAttribute("fld", filterField.Field);
                xml.WriteAttributeOptional("item", filterField.ItemIndex);
                xml.WriteAttributeOptional("hier", filterField.HierarchyIndex);
                xml.WriteAttributeOptional("name", filterField.HierarchyUniqueName);
                xml.WriteAttributeOptional("cap", filterField.HierarchyDisplayName);
                xml.WriteEndElement(); // pageField
            }

            xml.WriteEndElement(); // pageFields
        }

        if (pt.DataFields.Count > 0)
        {
            xml.WriteStartElement("dataFields", Main2006SsNs);
            xml.WriteAttribute("count", pt.DataFields.Count);
            foreach (XlsxPivotDataField dataField in pt.DataFields)
            {
                xml.WriteStartElement("dataField", Main2006SsNs);
                xml.WriteAttributeOptional("name", dataField.DataFieldName);
                xml.WriteAttribute("fld", dataField.Field);
                if (dataField.Subtotal != XlsxPivotSummary.Sum)
                {
                    var subtotalAttr = dataField.Subtotal switch
                    {
                        XlsxPivotSummary.Sum => "sum",
                        XlsxPivotSummary.Count => "count",
                        XlsxPivotSummary.Average => "average",
                        XlsxPivotSummary.Minimum => "min",
                        XlsxPivotSummary.Maximum => "max",
                        XlsxPivotSummary.Product => "product",
                        XlsxPivotSummary.CountNumbers => "countNums",
                        XlsxPivotSummary.StandardDeviation => "stdDev",
                        XlsxPivotSummary.PopulationStandardDeviation => "stdDevp",
                        XlsxPivotSummary.Variance => "var",
                        XlsxPivotSummary.PopulationVariance => "varp",
                        _ => throw new UnreachableException(),
                    };
                    xml.WriteAttribute("subtotal", subtotalAttr);
                }

                if (dataField.ShowDataAsFormat != XlsxPivotCalculation.Normal)
                {
                    var showDataAsAttr = dataField.ShowDataAsFormat switch
                    {
                        XlsxPivotCalculation.Normal => "normal",
                        XlsxPivotCalculation.DifferenceFrom => "difference",
                        XlsxPivotCalculation.PercentageOf => "percent",
                        XlsxPivotCalculation.PercentageDifferenceFrom => "percentDiff",
                        XlsxPivotCalculation.RunningTotal => "runTotal",
                        XlsxPivotCalculation.PercentageOfRow => "percentOfRow",
                        XlsxPivotCalculation.PercentageOfColumn => "percentOfCol",
                        XlsxPivotCalculation.PercentageOfTotal => "percentOfTotal",
                        XlsxPivotCalculation.Index => "index",
                        _ => throw new UnreachableException(),
                    };
                    xml.WriteAttribute("showDataAs", showDataAsAttr);
                }

                xml.WriteAttributeDefault("baseField", dataField.BaseField, -1);
                xml.WriteAttributeDefault("baseItem", dataField.BaseItem, 1048832);
                xml.WriteAttributeOptional("numFmtId", context.GetNumberFormatId(dataField.NumberFormatValue));

                xml.WriteEndElement(); // dataField
            }

            xml.WriteEndElement(); // dataFields
        }

        if (pt.Formats.Count > 0)
        {
            xml.WriteStartElement("formats", Main2006SsNs);
            xml.WriteAttribute("count", pt.Formats.Count);
            foreach (XlsxPivotFormat? format in pt.Formats)
            {
                xml.WriteStartElement("format", Main2006SsNs);
                if (format.Action != XlsxPivotFormatAction.Formatting)
                {
                    var actionAttr = format.Action switch
                    {
                        XlsxPivotFormatAction.Blank => "blank",
                        XlsxPivotFormatAction.Formatting => "formatting",
                        _ => throw new UnreachableException(),
                    };
                    xml.WriteAttribute("action", actionAttr);
                }

                // DxfId is technically optional.
                if (format.FormatValue is not null)
                {
                    var dxfId = context.GetDxfId(format.FormatValue);
                    xml.WriteAttribute("dxfId", dxfId);
                }

                XlsxPivotArea pivotArea = format.PivotArea;
                WritePivotArea(xml, pivotArea);
                xml.WriteEndElement(); // format
            }
            xml.WriteEndElement(); // formats
        }

        // Pivot table CF only specifies what should be formatted in PT. The actual CF
        // specification is in sheet:conditionalFormatting that with a flag Pivot="1"
        // and matching priority.
        if (pt.ConditionalFormats.Count > 0)
        {
            xml.WriteStartElement("conditionalFormats", Main2006SsNs);
            xml.WriteAttribute("count", pt.ConditionalFormats.Count);
            foreach (XlsxPivotConditionalFormat? conditionalFormat in pt.ConditionalFormats)
            {
                xml.WriteStartElement("conditionalFormat", Main2006SsNs);
                if (conditionalFormat.Scope != XlsxPivotCfScope.SelectedCells)
                {
                    var scopeAttr = conditionalFormat.Scope switch
                    {
                        XlsxPivotCfScope.SelectedCells => "selection",
                        XlsxPivotCfScope.DataFields => "data",
                        XlsxPivotCfScope.FieldIntersections => "field",
                        _ => throw new UnreachableException(),
                    };
                    xml.WriteAttribute("scope", scopeAttr);
                }

                if (conditionalFormat.Type != XlsxPivotCfRuleType.None)
                {
                    var typeAttr = conditionalFormat.Type switch
                    {
                        XlsxPivotCfRuleType.All => "all",
                        XlsxPivotCfRuleType.Column => "column",
                        XlsxPivotCfRuleType.None => "none",
                        XlsxPivotCfRuleType.Row => "row",
                        _ => throw new UnreachableException(),
                    };
                    xml.WriteAttribute("type", typeAttr);
                }

                xml.WriteAttribute("priority", conditionalFormat.Format.Priority);
                xml.WriteStartElement("pivotAreas", Main2006SsNs);
                xml.WriteAttribute("count", conditionalFormat.Areas.Count);
                foreach (XlsxPivotArea pivotArea in conditionalFormat.Areas)
                    WritePivotArea(xml, pivotArea);

                xml.WriteEndElement(); // pivotAreas
                xml.WriteEndElement(); // conditionalFormat
            }

            xml.WriteEndElement(); // conditionalFormats
        }

        var hasDefaultTheme =
            pt.Theme == XlsxPivotTableTheme.None &&
            pt.ShowRowHeaders == false &&
            pt.ShowColumnHeaders == false &&
            pt.ShowRowStripes == false &&
            pt.ShowColumnStripes == false &&
            pt.ShowLastColumn == false;
        if (!hasDefaultTheme)
        {
            xml.WriteStartElement("pivotTableStyleInfo", Main2006SsNs);
            if (pt.Theme != XlsxPivotTableTheme.None)
                xml.WriteAttribute("name", pt.Theme.ToString());

            xml.WriteAttributeDefault("showRowHeaders", pt.ShowRowHeaders, false);
            xml.WriteAttributeDefault("showColHeaders", pt.ShowColumnHeaders, false);
            xml.WriteAttributeDefault("showRowStripes", pt.ShowRowStripes, false);
            xml.WriteAttributeDefault("showColStripes", pt.ShowColumnStripes, false);
            xml.WriteAttributeDefault("showLastColumn", pt.ShowLastColumn, false);
            xml.WriteEndElement(); // pivotTableStyleInfo
        }

        // Because extensions are pretty large, always write them.
        xml.WriteStartElement("extLst");

        {
            // See [MS-XLSX] 2.2.4.5 Pivot Table
            xml.WriteStartElement("ext", Main2006SsNs);
            xml.WriteAttributeString("xmlns", "x14", null, X14Main2009SsNs);
            xml.WriteAttributeString("uri", "{962EF5D1-5CA2-4c93-8EF4-DBF5C05439D2}");
            xml.WriteStartElement("pivotTableDefinition", X14Main2009SsNs);
            xml.WriteAttribute("enableEdit", pt.EnableCellEditing);
            xml.WriteAttribute("hideValuesRow", !pt.ShowValuesRow);
            xml.WriteEndElement(); // pivotTableDefinition
            xml.WriteEndElement(); // ext
        }

        xml.WriteEndElement(); // extList

        xml.WriteEndElement(); // pivotTableDefinition

        xml.Close();
    }

    private static void WriteAxis(XmlWriter xml, XlsxPivotTableAxis axis, string fieldsElement, string itemsElement)
    {
        if (axis.Fields.Count > 0)
        {
            xml.WriteStartElement(fieldsElement, Main2006SsNs);
            xml.WriteAttribute("count", axis.Fields.Count);
            foreach (FieldIndex axisField in axis.Fields)
            {
                xml.WriteStartElement("field", Main2006SsNs);
                xml.WriteAttribute("x", axisField.Value);
                xml.WriteEndElement();
            }

            xml.WriteEndElement(); // rowFields
        }

        if (axis.Items.Count > 0)
        {
            xml.WriteStartElement(itemsElement, Main2006SsNs);
            xml.WriteAttribute("count", axis.Items.Count);

            IReadOnlyList<int> previous = [];
            foreach (XlsxPivotFieldAxisItem axisItem in axis.Items)
            {
                xml.WriteStartElement("i", Main2006SsNs);
                if (axisItem.ItemType != XlsxPivotItemType.Data)
                {
                    var itemTypeAttr = GetItemTypeAttr(axisItem.ItemType);
                    xml.WriteAttribute("t", itemTypeAttr);
                }

                // 'r' attribute means repeat data from previous axis item.
                var r = 0;
                var maxPrefixLen = Math.Min(previous.Count, axisItem.FieldItem.Count);
                while (r < maxPrefixLen && previous[r] == axisItem.FieldItem[r])
                    r++;

                // It seems that Excel always has at least one <x> element, not sure if necessary,
                // but it makes xml comparisons far easier. This is common for non-data type items.
                if (r > 0 && r == axisItem.FieldItem.Count)
                    r--;

                xml.WriteAttributeDefault("r", r, 0);
                xml.WriteAttributeDefault("i", axisItem.DataItem, 0); // Data field index

                foreach (var fieldItem in axisItem.FieldItem.Skip(r))
                {
                    xml.WriteStartElement("x", Main2006SsNs);
                    xml.WriteAttributeDefault("v", fieldItem, 0);
                    xml.WriteEndElement(); // x
                }

                xml.WriteEndElement(); // i
                previous = axisItem.FieldItem;
            }

            xml.WriteEndElement();
        }
    }

    private static void WritePivotArea(XmlWriter xml, XlsxPivotArea pivotArea)
    {
        xml.WriteStartElement("pivotArea", Main2006SsNs);
        xml.WriteAttributeOptional("field", pivotArea.Field?.Value);
        if (pivotArea.Type != XlsxPivotAreaType.Normal)
        {
            var typeAttr = pivotArea.Type switch
            {
                XlsxPivotAreaType.None => "none",
                XlsxPivotAreaType.Normal => "normal",
                XlsxPivotAreaType.Data => "data",
                XlsxPivotAreaType.All => "all",
                XlsxPivotAreaType.Origin => "origin",
                XlsxPivotAreaType.Button => "button",
                XlsxPivotAreaType.TopRight => "topRight",
                XlsxPivotAreaType.TopEnd => "topEnd",
                _ => throw new UnreachableException(),
            };
            xml.WriteAttribute("type", typeAttr);
        }

        xml.WriteAttributeDefault("dataOnly", pivotArea.DataOnly, true);
        xml.WriteAttributeDefault("labelOnly", pivotArea.LabelOnly, false);
        xml.WriteAttributeDefault("grandRow", pivotArea.GrandRow, false);
        xml.WriteAttributeDefault("grandCol", pivotArea.GrandCol, false);
        xml.WriteAttributeDefault("cacheIndex", pivotArea.CacheIndex, false);
        xml.WriteAttributeDefault("outline", pivotArea.Outline, true);
        if (pivotArea.Offset is not null)
            xml.WriteAttribute("offset", pivotArea.Offset.Value.ToString());

        xml.WriteAttributeDefault("collapsedLevelsAreSubtotals", pivotArea.CollapsedLevelsAreSubtotals, false);
        if (pivotArea.Axis is not null)
            xml.WriteAttribute("axis", GetAxisAttr(pivotArea.Axis.Value));

        xml.WriteAttributeOptional("fieldPosition", pivotArea.FieldPosition);

        if (pivotArea.References.Count > 0)
        {
            xml.WriteStartElement("references", Main2006SsNs);
            xml.WriteAttribute("count", pivotArea.References.Count);
            foreach (XlsxPivotReference reference in pivotArea.References)
            {
                xml.WriteStartElement("reference", Main2006SsNs);
                xml.WriteAttributeOptional("field", reference.Field);
                xml.WriteAttribute("count", reference.FieldItems.Count);
                xml.WriteAttributeDefault("selected", reference.Selected, true);
                xml.WriteAttributeDefault("byPosition", reference.ByPosition, false);
                xml.WriteAttributeDefault("relative", reference.Relative, false);
                HashSet<XlsxSubtotalFunction> subtotals = reference.Subtotals;
                xml.WriteAttributeDefault("defaultSubtotal", subtotals.Contains(XlsxSubtotalFunction.Automatic), false);
                xml.WriteAttributeDefault("sumSubtotal", subtotals.Contains(XlsxSubtotalFunction.Sum), false);
                xml.WriteAttributeDefault("countASubtotal", subtotals.Contains(XlsxSubtotalFunction.Count), false);
                xml.WriteAttributeDefault("avgSubtotal", subtotals.Contains(XlsxSubtotalFunction.Average), false);
                xml.WriteAttributeDefault("maxSubtotal", subtotals.Contains(XlsxSubtotalFunction.Maximum), false);
                xml.WriteAttributeDefault("minSubtotal", subtotals.Contains(XlsxSubtotalFunction.Minimum), false);
                xml.WriteAttributeDefault("productSubtotal", subtotals.Contains(XlsxSubtotalFunction.Product), false);
                xml.WriteAttributeDefault("countSubtotal", subtotals.Contains(XlsxSubtotalFunction.CountNumbers), false);
                xml.WriteAttributeDefault("stdDevSubtotal", subtotals.Contains(XlsxSubtotalFunction.StandardDeviation), false);
                xml.WriteAttributeDefault("stdDevPSubtotal", subtotals.Contains(XlsxSubtotalFunction.PopulationStandardDeviation), false);
                xml.WriteAttributeDefault("varSubtotal", subtotals.Contains(XlsxSubtotalFunction.Variance), false);
                xml.WriteAttributeDefault("varPSubtotal", subtotals.Contains(XlsxSubtotalFunction.PopulationVariance), false);

                foreach (var fieldItem in reference.FieldItems)
                {
                    xml.WriteStartElement("x", Main2006SsNs);
                    xml.WriteAttribute("v", fieldItem);
                    xml.WriteEndElement(); // x
                }

                xml.WriteEndElement(); // reference
            }

            xml.WriteEndElement(); // references
        }

        xml.WriteEndElement(); // pivotArea
    }

    private static string GetItemTypeAttr(XlsxPivotItemType itemType)
    {
        var itemTypeAttr = itemType switch
        {
            XlsxPivotItemType.Avg => "avg",
            XlsxPivotItemType.Blank => "blank",
            XlsxPivotItemType.Count => "count",
            XlsxPivotItemType.CountA => "countA",
            XlsxPivotItemType.Data => "data",
            XlsxPivotItemType.Default => "default",
            XlsxPivotItemType.Grand => "grand",
            XlsxPivotItemType.Max => "max",
            XlsxPivotItemType.Min => "min",
            XlsxPivotItemType.Product => "product",
            XlsxPivotItemType.StdDev => "stdDev",
            XlsxPivotItemType.StdDevP => "stdDevP",
            XlsxPivotItemType.Sum => "sum",
            XlsxPivotItemType.Var => "var",
            XlsxPivotItemType.VarP => "varP",
            _ => throw new UnreachableException(),
        };
        return itemTypeAttr;
    }

    private static string GetAxisAttr(XlsxPivotAxis axis)
    {
        return axis switch
        {
            XlsxPivotAxis.AxisRow => "axisRow",
            XlsxPivotAxis.AxisCol => "axisCol",
            XlsxPivotAxis.AxisPage => "axisPage",
            XlsxPivotAxis.AxisValues => "axisValues",
            _ => throw new UnreachableException(),
        };
    }
}
