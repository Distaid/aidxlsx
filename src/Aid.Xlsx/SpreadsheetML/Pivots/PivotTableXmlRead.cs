using System.Xml;
using System.Xml.Linq;
using Aid.Xlsx.Model.ConditionalFormats;
using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.Grid.Coordinates;
using Aid.Xlsx.Model.PivotTables;
using Aid.Xlsx.Model.PivotTables.Areas;
using Aid.Xlsx.Model.PivotTables.PivotValues;
using Aid.Xlsx.Model.Styles;
using Aid.Xlsx.Model.Styles.Formatting;
using Aid.Xlsx.Model.Worksheets;
using Aid.Xlsx.Opc.Xml;
using Aid.Xlsx.SpreadsheetML.Workbooks;
using Aid.Xlsx.SpreadsheetML.Xml;

namespace Aid.Xlsx.SpreadsheetML.Pivots;

internal static partial class PivotTableXml
{
    private const int ValuesFieldIndex = -2;

    internal static XElement Read(Stream stream)
    {
        using var reader = XmlReader.Create(stream, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit });
        var root = XElement.Load(reader, System.Xml.Linq.LoadOptions.PreserveWhitespace);
        if (root.Name.LocalName != "pivotTableDefinition" || root.Name.NamespaceName is not (OpenXmlConst.Main2006SsNs or OpenXmlConst.StrictSpreadsheetNs))
            throw PartStructureException.IncorrectElementFormat("pivotTableDefinition");

        return root;
    }

    internal static void Load(XElement definition, XlsxWorksheet sheet, XlsxPivotCache cache, string relId, string cacheRelId, LoadContext context)
    {
        var pivot = LoadPivotTableDefinition(definition, sheet, cache, context);
        var area = pivot.Area;
        sheet.Range(area.TopRow, area.LeftColumn, area.BottomRow, area.RightColumn).Clear(XlsxClearOptions.All);
        sheet.PivotTables.Add(pivot);
        pivot.RelId = relId;
        pivot.CacheDefinitionRelId = cacheRelId;
    }

    private static XElement? Child(XElement? parent, string name) => parent?.Element(parent.Name.Namespace + name);

    private static IEnumerable<XElement> Children(XElement parent, string name) => parent.Elements(parent.Name.Namespace + name);

    private static XElement? Extension(XElement parent, string name) =>
        Child(parent, "extLst")?.Elements(parent.Name.Namespace + "ext")
            .SelectMany(ext => ext.Elements(XName.Get(name, OpenXmlConst.X14Main2009SsNs))).FirstOrDefault();

    private static XlsxPivotTable LoadPivotTableDefinition(XElement pivotTable, XlsxWorksheet sheet, XlsxPivotCache cache, LoadContext context)
    {
        var styles = sheet.Workbook.Styles;

        // Load base attributes
        var xlPivotTable = LoadPivotTableAttributes(pivotTable, sheet, cache);

        // Load location
        var location = Child(pivotTable, "location") ?? throw PartStructureException.ExpectedElementNotFound();
        var referenceText = (string?)location.Attribute("ref") ?? throw PartStructureException.MissingAttribute();
        xlPivotTable.Area = Area.Parse(referenceText);
        xlPivotTable.FirstHeaderRow = (uint?)location.Attribute("firstHeaderRow") ?? throw PartStructureException.MissingAttribute();
        xlPivotTable.FirstDataRow = (uint?)location.Attribute("firstDataRow") ?? throw PartStructureException.MissingAttribute();
        xlPivotTable.FirstDataCol = (uint?)location.Attribute("firstDataCol") ?? throw PartStructureException.MissingAttribute();

        // Skip `rowPageCount` and `colPageCount`, because they are derived from filterAreaOrder, filterFieldsPageWrap and pageField count

        // Load pivot fields
        var pivotFields = Child(pivotTable, "pivotFields");
        if (pivotFields is not null)
        {
            foreach (var pivotField in Children(pivotFields, "pivotField"))
                xlPivotTable.AddField(LoadPivotField(pivotField, xlPivotTable, styles));
        }

        // Load row axis fields and items
        LoadAxisFields(Child(pivotTable, "rowFields"), xlPivotTable.RowAxis, xlPivotTable);
        LoadAxisItems(Child(pivotTable, "rowItems"), xlPivotTable.RowAxis);

        // Load column axis fields and items
        LoadAxisFields(Child(pivotTable, "colFields"), xlPivotTable.ColumnAxis, xlPivotTable);
        LoadAxisItems(Child(pivotTable, "colItems"), xlPivotTable.ColumnAxis);

        // Load page fields, i.e. the filters region.
        var pageFields = Child(pivotTable, "pageFields");
        if (pageFields is not null)
        {
            foreach (var pageField in Children(pageFields, "pageField"))
            {
                var field = (int?)pageField.Attribute("fld") ?? throw PartStructureException.MissingAttribute();
                var itemIndex = checked((int?)(uint?)pageField.Attribute("item"));
                var hierarchyIndex = (int?)pageField.Attribute("hier");
                var hierarchyUniqueName = (string?)pageField.Attribute("name");
                var hierarchyDisplayName = (string?)pageField.Attribute("cap");
                var xlPageField = new XlsxPivotPageField(field)
                {
                    ItemIndex = itemIndex,
                    HierarchyIndex = hierarchyIndex,
                    HierarchyUniqueName = hierarchyUniqueName,
                    HierarchyDisplayName = hierarchyDisplayName,
                };
                xlPivotTable.Filters.AddField(xlPageField);
            }
        }

        // Load data fields.
        var dataFields = Child(pivotTable, "dataFields");
        if (dataFields is not null)
        {
            foreach (var dataField in Children(dataFields, "dataField"))
            {
                var name = (string?)dataField.Attribute("name");
                var field = (uint?)dataField.Attribute("fld") ?? throw PartStructureException.MissingAttribute();
                var subtotal = ParseSummary((string?)dataField.Attribute("subtotal")) ?? XlsxPivotSummary.Sum;
                var showDataAsFormat = ParseCalculation((string?)dataField.Attribute("showDataAs")) ?? XlsxPivotCalculation.Normal;
                var baseField = (int?)dataField.Attribute("baseField") ?? -1;
                var baseItem = (uint?)dataField.Attribute("baseItem") ?? 1048832;
                var numberFormatId = checked((int?)(uint?)dataField.Attribute("numFmtId"));
                var numberFormat = numberFormatId is not null ? styles.NumberFormats[numberFormatId.Value] : (XlsxNumberFormat?)null;
                var xlDataField = new XlsxPivotDataField(xlPivotTable, checked((int)field))
                {
                    DataFieldName = name,
                    Subtotal = subtotal,
                    ShowDataAsFormat = showDataAsFormat,
                    BaseField = baseField,
                    BaseItem = baseItem,
                    NumberFormatValue = numberFormat,
                };
                xlPivotTable.DataFields.AddField(xlDataField);
            }
        }

        // Load formats
        var formats = Child(pivotTable, "formats");
        if (formats is not null)
        {
            foreach (var format in Children(formats, "format"))
            {
                var action = ParseFormatAction((string?)format.Attribute("action")) ?? XlsxPivotFormatAction.Formatting;
                var dxf = (uint?)format.Attribute("dxfId") is { } dxfId
                    ? sheet.Workbook.Styles.DifferentialFormats[checked((int)dxfId)]
                    : null;
                var pivotArea = Child(format, "pivotArea") ?? throw PartStructureException.ExpectedElementNotFound();
                var xlPivotArea = LoadPivotArea(pivotArea);
                var xlFormat = new XlsxPivotFormat(xlPivotArea)
                {
                    Action = action,
                    FormatValue = dxf,
                };
                xlPivotTable.AddFormat(xlFormat);
            }
        }

        var conditionalFormats = Child(pivotTable, "conditionalFormats");
        if (conditionalFormats is not null)
        {
            foreach (var conditionalFormat in Children(conditionalFormats, "conditionalFormat"))
            {
                var scope = ParseCfScope((string?)conditionalFormat.Attribute("scope")) ?? XlsxPivotCfScope.SelectedCells;
                var type = ParseCfRuleType((string?)conditionalFormat.Attribute("type")) ?? XlsxPivotCfRuleType.None;
                var priority = (uint?)conditionalFormat.Attribute("priority") ?? throw PartStructureException.MissingAttribute();
                var format = context.GetPivotCf(sheet.Name, checked((int)priority));
                var xlConditionalFormat = new XlsxPivotConditionalFormat(format)
                {
                    Scope = scope,
                    Type = type,
                };
                var pivotAreas = Child(conditionalFormat, "pivotAreas");
                if (pivotAreas is not null)
                {
                    foreach (var pivotArea in Children(pivotAreas, "pivotArea"))
                    {
                        var xlPivotArea = LoadPivotArea(pivotArea);
                        xlConditionalFormat.AddArea(xlPivotArea);
                    }
                }

                xlPivotTable.AddConditionalFormat(xlConditionalFormat);
            }
        }

        // TODO: chartFormats
        // pivotHierarchies is OLAP and thus for now out of scope.
        var pivotTableStyle = Child(pivotTable, "pivotTableStyleInfo");
        LoadPivotTableStyle(pivotTableStyle, xlPivotTable);

        // TODO: filters
        // rowHierarchiesUsage is OLAP and thus for now out of scope.
        // colHierarchiesUsage is OLAP and thus for now out of scope.
        LoadExtensionList(pivotTable, xlPivotTable);

        return xlPivotTable;
    }

    private static XlsxPivotTable LoadPivotTableAttributes(XElement pivotTable, XlsxWorksheet sheet, XlsxPivotCache cache)
    {
        var name = (string?)pivotTable.Attribute("name") ?? throw PartStructureException.MissingAttribute();
        _ = (uint?)pivotTable.Attribute("cacheId") ?? throw PartStructureException.MissingAttribute("cacheId");
        var dataOnRows = (bool?)pivotTable.Attribute("dataOnRows") ?? false;

        // DataPosition attribute is skipped, because it basically represents a field on one of axis.
        // Excel requires that dataPosition and field with index -2 must be in list of respective axis
        // at correct place, otherwise it crashes. To make things simple, we set the value when it is
        // encountered on the correct axis (plus there is a check that field is not used on multiple axes
        // that would cause exception).
        var autoFormatId = (uint?)pivotTable.Attribute("autoFormatId");
        var applyNumberFormats = (bool?)pivotTable.Attribute("applyNumberFormats") ?? false;
        var applyBorderFormats = (bool?)pivotTable.Attribute("applyBorderFormats") ?? false;
        var applyFontFormats = (bool?)pivotTable.Attribute("applyFontFormats") ?? false;
        var applyPatternFormats = (bool?)pivotTable.Attribute("applyPatternFormats") ?? false;
        var applyAlignmentFormats = (bool?)pivotTable.Attribute("applyAlignmentFormats") ?? false;
        var applyWidthHeightFormats = (bool?)pivotTable.Attribute("applyWidthHeightFormats") ?? false;
        var dataCaption = (string?)pivotTable.Attribute("dataCaption") ?? throw PartStructureException.MissingAttribute();
        var grandTotalCaption = (string?)pivotTable.Attribute("grandTotalCaption");
        var errorCaption = (string?)pivotTable.Attribute("errorCaption");
        var showError = (bool?)pivotTable.Attribute("showError") ?? false;
        var missingCaption = (string?)pivotTable.Attribute("missingCaption") ?? string.Empty;
        var showMissing = (bool?)pivotTable.Attribute("showMissing") ?? true;
        var pageStyle = (string?)pivotTable.Attribute("pageStyle");
        var pivotTableStyleName = (string?)pivotTable.Attribute("pivotTableStyle");
        var vacatedStyle = (string?)pivotTable.Attribute("vacatedStyle");
        var tag = (string?)pivotTable.Attribute("tag");
        var updatedVersion = checked((byte)((uint?)pivotTable.Attribute("updatedVersion") ?? 0));
        var minRefreshableVersion = checked((byte)((uint?)pivotTable.Attribute("minRefreshableVersion") ?? 0));
        var asteriskTotals = (bool?)pivotTable.Attribute("asteriskTotals") ?? false;
        var showItems = (bool?)pivotTable.Attribute("showItems") ?? true;
        var editData = (bool?)pivotTable.Attribute("editData") ?? false;
        var disableFieldList = (bool?)pivotTable.Attribute("disableFieldList") ?? false;
        var showCalculatedMembers = (bool?)pivotTable.Attribute("showCalcMbrs") ?? true;
        var visualTotals = (bool?)pivotTable.Attribute("visualTotals") ?? true;
        var showMultipleLabel = (bool?)pivotTable.Attribute("showMultipleLabel") ?? true;
        var showDataDropDown = (bool?)pivotTable.Attribute("showDataDropDown") ?? true;
        var showDrill = (bool?)pivotTable.Attribute("showDrill") ?? true;
        var printDrill = (bool?)pivotTable.Attribute("printDrill") ?? false;
        var showMemberPropertyTips = (bool?)pivotTable.Attribute("showMemberPropertyTips") ?? true;
        var showDataTips = (bool?)pivotTable.Attribute("showDataTips") ?? true;
        var enableWizard = (bool?)pivotTable.Attribute("enableWizard") ?? true;
        var enableDrill = (bool?)pivotTable.Attribute("enableDrill") ?? true;
        var enableFieldProperties = (bool?)pivotTable.Attribute("enableFieldProperties") ?? true;
        var preserveFormatting = (bool?)pivotTable.Attribute("preserveFormatting") ?? true;
        var useAutoFormatting = (bool?)pivotTable.Attribute("useAutoFormatting") ?? false;
        var pageWrap = (uint?)pivotTable.Attribute("pageWrap") ?? 0;
        var pageOverThenDown = (bool?)pivotTable.Attribute("pageOverThenDown") ?? false;
        var subtotalHiddenItems = (bool?)pivotTable.Attribute("subtotalHiddenItems") ?? false;
        var rowGrandTotals = (bool?)pivotTable.Attribute("rowGrandTotals") ?? true;
        var columnGrandTotals = (bool?)pivotTable.Attribute("colGrandTotals") ?? true;
        var fieldPrintTitles = (bool?)pivotTable.Attribute("fieldPrintTitles") ?? false;
        var itemPrintTitles = (bool?)pivotTable.Attribute("itemPrintTitles") ?? false;
        var mergeItem = (bool?)pivotTable.Attribute("mergeItem") ?? false;
        var showDropZones = (bool?)pivotTable.Attribute("showDropZones") ?? true;
        var createdVersion = checked((byte)((uint?)pivotTable.Attribute("createdVersion") ?? 0));
        var indent = (uint?)pivotTable.Attribute("indent") ?? 1;
        var showEmptyRow = (bool?)pivotTable.Attribute("showEmptyRow") ?? false;
        var showEmptyColumn = (bool?)pivotTable.Attribute("showEmptyCol") ?? false;
        var showHeaders = (bool?)pivotTable.Attribute("showHeaders") ?? true;
        var compact = (bool?)pivotTable.Attribute("compact") ?? true;
        var outline = (bool?)pivotTable.Attribute("outline") ?? false;
        var outlineData = (bool?)pivotTable.Attribute("outlineData") ?? false;
        var compactData = (bool?)pivotTable.Attribute("compactData") ?? true;
        var published = (bool?)pivotTable.Attribute("published") ?? false;
        var gridDropZones = (bool?)pivotTable.Attribute("gridDropZones") ?? false;
        var stopImmersiveUi = (bool?)pivotTable.Attribute("immersive") ?? true;
        var multipleFieldFilters = (bool?)pivotTable.Attribute("multipleFieldFilters") ?? true;
        var chartFormat = (uint?)pivotTable.Attribute("chartFormat") ?? 0;
        var rowHeaderCaption = (string?)pivotTable.Attribute("rowHeaderCaption");
        var columnHeaderCaption = (string?)pivotTable.Attribute("colHeaderCaption");
        var fieldListSortAscending = (bool?)pivotTable.Attribute("fieldListSortAscending") ?? false;
        var mdxSubQueries = (bool?)pivotTable.Attribute("mdxSubqueries") ?? false;
        var customSortList = (bool?)pivotTable.Attribute("customListSort") ?? true;

        var xlPivotTable = new XlsxPivotTable(sheet, cache)
        {
            Name = name,
            DataOnRows = dataOnRows,
            DataPosition = null, // 'data' field is set when during axis loading (if present).
            AutoFormatId = autoFormatId,
            ApplyNumberFormats = applyNumberFormats,
            ApplyBorderFormats = applyBorderFormats,
            ApplyFontFormats = applyFontFormats,
            ApplyPatternFormats = applyPatternFormats,
            ApplyAlignmentFormats = applyAlignmentFormats,
            ApplyWidthHeightFormats = applyWidthHeightFormats,
            DataCaption = dataCaption,
            GrandTotalCaption = grandTotalCaption,
            ErrorValueReplacement = errorCaption,
            ShowError = showError,
            MissingCaption = missingCaption,
            ShowMissing = showMissing,
            PageStyle = pageStyle,
            PivotTableStyleName = pivotTableStyleName,
            VacatedStyle = vacatedStyle,
            Tag = tag,
            UpdatedVersion = updatedVersion,
            MinRefreshableVersion = minRefreshableVersion,
            AsteriskTotals = asteriskTotals,
            DisplayItemLabels = showItems,
            EditData = editData,
            DisableFieldList = disableFieldList,
            ShowCalculatedMembers = showCalculatedMembers,
            VisualTotals = visualTotals,
            ShowMultipleLabel = showMultipleLabel,
            ShowDataDropDown = showDataDropDown,
            ShowExpandCollapseButtons = showDrill,
            PrintExpandCollapsedButtons = printDrill,
            ShowPropertiesInTooltips = showMemberPropertyTips,
            ShowContextualTooltips = showDataTips,
            EnableEditingMechanism = enableWizard,
            EnableShowDetails = enableDrill,
            EnableFieldProperties = enableFieldProperties,
            PreserveCellFormatting = preserveFormatting,
            AutofitColumns = useAutoFormatting,
            FilterFieldsPageWrap = checked((int)pageWrap),
            FilterAreaOrder = pageOverThenDown ? XlsxFilterAreaOrder.OverThenDown : XlsxFilterAreaOrder.DownThenOver,
            FilteredItemsInSubtotals = subtotalHiddenItems,
            ShowGrandTotalsRows = rowGrandTotals,
            ShowGrandTotalsColumns = columnGrandTotals,
            PrintTitles = fieldPrintTitles,
            RepeatRowLabels = itemPrintTitles,
            MergeAndCenterWithLabels = mergeItem,
            ShowDropZones = showDropZones,
            PivotCacheCreatedVersion = createdVersion,
            RowLabelIndent = checked((int)indent),
            ShowEmptyItemsOnRows = showEmptyRow,
            ShowEmptyItemsOnColumns = showEmptyColumn,
            DisplayCaptionsAndDropdowns = showHeaders,
            Compact = compact,
            Outline = outline,
            OutlineData = outlineData,
            CompactData = compactData,
            Published = published,
            ClassicPivotTableLayout = gridDropZones,
            StopImmersiveUi = stopImmersiveUi,
            AllowMultipleFilters = multipleFieldFilters,
            ChartFormat = chartFormat,
            RowHeaderCaption = rowHeaderCaption,
            ColumnHeaderCaption = columnHeaderCaption,
            SortFieldsAtoZ = fieldListSortAscending,
            MdxSubQueries = mdxSubQueries,
            UseCustomListsForSorting = customSortList,
        };
        return xlPivotTable;
    }

    private static XlsxPivotTableField LoadPivotField(XElement pivotField, XlsxPivotTable xlPivotTable, XlsxWorkbookStyles styles)
    {
        var customName = (string?)pivotField.Attribute("name");
        var axis = ParseAxis((string?)pivotField.Attribute("axis"));
        var dataField = (bool?)pivotField.Attribute("dataField") ?? false;
        var subtotalCaption = (string?)pivotField.Attribute("subtotalCaption");
        var showDropDowns = (bool?)pivotField.Attribute("showDropDowns") ?? true;
        var hiddenLevel = (bool?)pivotField.Attribute("hiddenLevel") ?? false;
        var uniqueMemberProperty = (string?)pivotField.Attribute("uniqueMemberProperty");
        var compact = (bool?)pivotField.Attribute("compact") ?? true;
        var allDrilled = (bool?)pivotField.Attribute("allDrilled") ?? false;
        var numberFormatId = checked((int?)(uint?)pivotField.Attribute("numFmtId"));
        var numberFormat = numberFormatId is not null ? styles.NumberFormats[numberFormatId.Value] : (XlsxNumberFormat?)null;
        var outline = (bool?)pivotField.Attribute("outline") ?? true;
        var subtotalTop = (bool?)pivotField.Attribute("subtotalTop") ?? true;
        var dragToRow = (bool?)pivotField.Attribute("dragToRow") ?? true;
        var dragToColumn = (bool?)pivotField.Attribute("dragToCol") ?? true;
        var multipleItemSelectionAllowed = (bool?)pivotField.Attribute("multipleItemSelectionAllowed") ?? false;
        var dragToPage = (bool?)pivotField.Attribute("dragToPage") ?? true;
        var dragToData = (bool?)pivotField.Attribute("dragToData") ?? true;
        var dragOff = (bool?)pivotField.Attribute("dragOff") ?? true;
        var showAll = (bool?)pivotField.Attribute("showAll") ?? true;
        var insertBlankRow = (bool?)pivotField.Attribute("insertBlankRow") ?? false;
        var serverField = (bool?)pivotField.Attribute("serverField") ?? false;
        var insertPageBreak = (bool?)pivotField.Attribute("insertPageBreak") ?? false;
        var autoShow = (bool?)pivotField.Attribute("autoShow") ?? false;
        var topAutoShow = (bool?)pivotField.Attribute("topAutoShow") ?? true;
        var hideNewItems = (bool?)pivotField.Attribute("hideNewItems") ?? false;
        var measureFilter = (bool?)pivotField.Attribute("measureFilter") ?? false;
        var includeNewItemsInFilter = (bool?)pivotField.Attribute("includeNewItemsInFilter") ?? false;
        var itemPageCount = (uint?)pivotField.Attribute("itemPageCount") ?? 10u;
        var sortType = ParseSortType((string?)pivotField.Attribute("sortType")) ?? XlsxPivotSortType.Default;
        var dataSourceSort = (bool?)pivotField.Attribute("dataSourceSort");
        var nonAutoSortDefault = (bool?)pivotField.Attribute("nonAutoSortDefault") ?? false;
        var rankBy = (uint?)pivotField.Attribute("rankBy");
        var defaultSubtotal = (bool?)pivotField.Attribute("defaultSubtotal") ?? true;
        var sumSubtotal = (bool?)pivotField.Attribute("sumSubtotal") ?? false;
        var countASubtotal = (bool?)pivotField.Attribute("countASubtotal") ?? false;
        var avgSubtotal = (bool?)pivotField.Attribute("avgSubtotal") ?? false;
        var maxSubtotal = (bool?)pivotField.Attribute("maxSubtotal") ?? false;
        var minSubtotal = (bool?)pivotField.Attribute("minSubtotal") ?? false;
        var productSubtotal = (bool?)pivotField.Attribute("productSubtotal") ?? false;
        var countSubtotal = (bool?)pivotField.Attribute("countSubtotal") ?? false;
        var stdDevSubtotal = (bool?)pivotField.Attribute("stdDevSubtotal") ?? false;
        var stdDevPSubtotal = (bool?)pivotField.Attribute("stdDevPSubtotal") ?? false;
        var varSubtotal = (bool?)pivotField.Attribute("varSubtotal") ?? false;
        var varPSubtotal = (bool?)pivotField.Attribute("varPSubtotal") ?? false;
        var showPropCell = (bool?)pivotField.Attribute("showPropCell") ?? false;
        var showPropTip = (bool?)pivotField.Attribute("showPropTip") ?? false;
        var showPropAsCaption = (bool?)pivotField.Attribute("showPropAsCaption") ?? false;
        var defaultAttributeDrillState = (bool?)pivotField.Attribute("defaultAttributeDrillState") ?? false;

        var subtotals = new HashSet<XlsxSubtotalFunction>();
        if (defaultSubtotal)
            subtotals.Add(XlsxSubtotalFunction.Automatic);

        if (sumSubtotal)
            subtotals.Add(XlsxSubtotalFunction.Sum);

        if (countASubtotal)
            subtotals.Add(XlsxSubtotalFunction.Count);

        if (avgSubtotal)
            subtotals.Add(XlsxSubtotalFunction.Average);

        if (maxSubtotal)
            subtotals.Add(XlsxSubtotalFunction.Maximum);

        if (minSubtotal)
            subtotals.Add(XlsxSubtotalFunction.Minimum);

        if (productSubtotal)
            subtotals.Add(XlsxSubtotalFunction.Product);

        if (countSubtotal)
            subtotals.Add(XlsxSubtotalFunction.CountNumbers);

        if (stdDevSubtotal)
            subtotals.Add(XlsxSubtotalFunction.StandardDeviation);

        if (stdDevPSubtotal)
            subtotals.Add(XlsxSubtotalFunction.PopulationStandardDeviation);

        if (varSubtotal)
            subtotals.Add(XlsxSubtotalFunction.Variance);

        if (varPSubtotal)
            subtotals.Add(XlsxSubtotalFunction.PopulationVariance);

        var xlField = new XlsxPivotTableField(xlPivotTable)
        {
            Name = customName,
            Axis = axis,
            DataField = dataField,
            SubtotalCaption = subtotalCaption ?? string.Empty,
            ShowDropDowns = showDropDowns,
            HiddenLevel = hiddenLevel,
            UniqueMemberProperty = uniqueMemberProperty,
            Compact = compact,
            AllDrilled = allDrilled,
            NumberFormatValue = numberFormat,
            Outline = outline,
            SubtotalTop = subtotalTop,
            DragToRow = dragToRow,
            DragToColumn = dragToColumn,
            MultipleItemSelectionAllowed = multipleItemSelectionAllowed,
            DragToPage = dragToPage,
            DragToData = dragToData,
            DragOff = dragOff,
            ShowAll = showAll,
            InsertBlankRow = insertBlankRow,
            ServerField = serverField,
            InsertPageBreak = insertPageBreak,
            AutoShow = autoShow,
            TopAutoShow = topAutoShow,
            HideNewItems = hideNewItems,
            MeasureFilter = measureFilter,
            IncludeNewItemsInFilter = includeNewItemsInFilter,
            ItemPageCount = itemPageCount,
            SortType = sortType,
            DataSourceSort = dataSourceSort,
            NonAutoSortDefault = nonAutoSortDefault,
            RankBy = rankBy,
            Subtotals = subtotals,
            ShowPropCell = showPropCell,
            ShowPropTip = showPropTip,
            ShowPropAsCaption = showPropAsCaption,
            DefaultAttributeDrillState = defaultAttributeDrillState,
        };

        var items = Child(pivotField, "items");
        if (items is not null)
        {
            foreach (var item in Children(items, "item"))
            {
                // Attributes `sd` and `d` were swapped in spec.
                var approximatelyHasChildren = (bool?)item.Attribute("c") ?? false;
                var details = (bool?)item.Attribute("d") ?? false;
                var drillAcrossAttributes = (bool?)item.Attribute("e") ?? true;
                var calculatedMember = (bool?)item.Attribute("f") ?? false;
                var hidden = (bool?)item.Attribute("h") ?? false;
                var missing = (bool?)item.Attribute("m") ?? false;
                var itemUserCaption = (string?)item.Attribute("n");
                var valueIsString = (bool?)item.Attribute("s") ?? false;
                var showDetails = (bool?)item.Attribute("sd") ?? true;
                var itemIndex = (uint?)item.Attribute("x");
                var itemType = ParseItemType((string?)item.Attribute("t")) ?? XlsxPivotItemType.Data;
                var xlItem = new XlsxPivotFieldItem(xlField, itemIndex is null ? null : checked((int)itemIndex.Value))
                {
                    ApproximatelyHasChildren = approximatelyHasChildren,
                    Details = details,
                    DrillAcrossAttributes = drillAcrossAttributes,
                    CalculatedMember = calculatedMember,
                    Hidden = hidden,
                    Missing = missing,
                    ItemUserCaption = itemUserCaption,
                    ValueIsString = valueIsString,
                    ShowDetails = showDetails,
                    ItemType = itemType,
                };

                xlField.AddItem(xlItem);
            }
        }

        // TODO: autoSortScope

        // extLst
        var field2010 = Extension(pivotField, "pivotField");
        xlField.RepeatItemLabels = (bool?)field2010?.Attribute("fillDownLabels") ?? false;

        return xlField;
    }

    private static void LoadAxisFields(XElement? fields, XlsxPivotTableAxis axis, XlsxPivotTable xlPivotTable)
    {
        if (fields is not null)
        {
            foreach (var field in Children(fields, "field"))
            {
                // Axis can contain 'data' field.
                var fieldIndex = (int?)field.Attribute("x") ?? throw PartStructureException.MissingAttribute();
                if (fieldIndex >= xlPivotTable.PivotFields.Count || (fieldIndex < 0 && fieldIndex != ValuesFieldIndex))
                    throw PartStructureException.InvalidAttributeValue();

                axis.AddField(fieldIndex);
            }
        }
    }

    private static void LoadAxisItems(XElement? axisItems, XlsxPivotTableAxis axis)
    {
        if (axisItems is not null)
        {
            // Both row and column use RowItem type for axis item.
            var previous = new List<int>();
            foreach (var axisItem in Children(axisItems, "i"))
            {
                var xlItemType = ParseItemType((string?)axisItem.Attribute("t")) ?? XlsxPivotItemType.Data;
                var dataFieldIndex = checked((int)((uint?)axisItem.Attribute("i") ?? 0)); // This is used by 'data' field
                var repeatedCount = (uint?)axisItem.Attribute("r") ?? 0;
                var fieldIndexes = new List<int>();
                foreach (var dataIndex in Children(axisItem, "x"))
                    fieldIndexes.Add((int?)dataIndex.Attribute("v") ?? 0);

                var allFieldIndexes = previous.Take((int)repeatedCount).Concat(fieldIndexes).ToList();
                axis.AddItem(new XlsxPivotFieldAxisItem(xlItemType, dataFieldIndex, allFieldIndexes));
                previous = allFieldIndexes;
            }
        }
    }

    private static XlsxPivotArea LoadPivotArea(XElement pivotArea)
    {
        var field = (int?)pivotArea.Attribute("field");
        var type = ParseAreaType((string?)pivotArea.Attribute("type")) ?? XlsxPivotAreaType.Normal;
        var dataOnly = (bool?)pivotArea.Attribute("dataOnly") ?? true;
        var labelOnly = (bool?)pivotArea.Attribute("labelOnly") ?? false;
        var grandRow = (bool?)pivotArea.Attribute("grandRow") ?? false;
        var grandCol = (bool?)pivotArea.Attribute("grandCol") ?? false;
        var cacheIndex = (bool?)pivotArea.Attribute("cacheIndex") ?? false;
        var outline = (bool?)pivotArea.Attribute("outline") ?? true;
        var offset = (string?)pivotArea.Attribute("offset") is { } offsetRefText ? Area.Parse(offsetRefText) : (Area?)null;
        var collapsedLevelsAreSubtotals = (bool?)pivotArea.Attribute("collapsedLevelsAreSubtotals") ?? false;
        var axis = ParseAxis((string?)pivotArea.Attribute("axis"));
        var fieldPosition = (uint?)pivotArea.Attribute("fieldPosition");
        var xlPivotArea = new XlsxPivotArea
        {
            Field = field,
            Type = type,
            DataOnly = dataOnly,
            LabelOnly = labelOnly,
            GrandRow = grandRow,
            GrandCol = grandCol,
            CacheIndex = cacheIndex,
            Outline = outline,
            Offset = offset,
            CollapsedLevelsAreSubtotals = collapsedLevelsAreSubtotals,
            Axis = axis,
            FieldPosition = fieldPosition
        };

        // Can contain extensions, in theory at least.
        var references = Child(pivotArea, "references");
        if (references is not null)
        {
            foreach (var reference in Children(references, "reference"))
                xlPivotArea.AddReference(LoadPivotReference(reference));
        }

        return xlPivotArea;
    }

    private static XlsxPivotReference LoadPivotReference(XElement reference)
    {
        var field = (uint?)reference.Attribute("field");
        var selected = (bool?)reference.Attribute("selected") ?? true;
        var byPosition = (bool?)reference.Attribute("byPosition") ?? false;
        var relative = (bool?)reference.Attribute("relative") ?? false;
        var defaultSubtotal = (bool?)reference.Attribute("defaultSubtotal") ?? false;
        var sumSubtotal = (bool?)reference.Attribute("sumSubtotal") ?? false;
        var countASubtotal = (bool?)reference.Attribute("countASubtotal") ?? false;
        var avgSubtotal = (bool?)reference.Attribute("avgSubtotal") ?? false;
        var maxSubtotal = (bool?)reference.Attribute("maxSubtotal") ?? false;
        var minSubtotal = (bool?)reference.Attribute("minSubtotal") ?? false;
        var productSubtotal = (bool?)reference.Attribute("productSubtotal") ?? false;
        var countSubtotal = (bool?)reference.Attribute("countSubtotal") ?? false;
        var stdDevSubtotal = (bool?)reference.Attribute("stdDevSubtotal") ?? false;
        var stdDevPSubtotal = (bool?)reference.Attribute("stdDevPSubtotal") ?? false;
        var varSubtotal = (bool?)reference.Attribute("varSubtotal") ?? false;
        var varPSubtotal = (bool?)reference.Attribute("varPSubtotal") ?? false;

        var subtotals = new HashSet<XlsxSubtotalFunction>();
        if (defaultSubtotal)
            subtotals.Add(XlsxSubtotalFunction.Automatic);

        if (sumSubtotal)
            subtotals.Add(XlsxSubtotalFunction.Sum);

        if (countASubtotal)
            subtotals.Add(XlsxSubtotalFunction.Count);

        if (avgSubtotal)
            subtotals.Add(XlsxSubtotalFunction.Average);

        if (maxSubtotal)
            subtotals.Add(XlsxSubtotalFunction.Maximum);

        if (minSubtotal)
            subtotals.Add(XlsxSubtotalFunction.Minimum);

        if (productSubtotal)
            subtotals.Add(XlsxSubtotalFunction.Product);

        if (countSubtotal)
            subtotals.Add(XlsxSubtotalFunction.CountNumbers);

        if (stdDevSubtotal)
            subtotals.Add(XlsxSubtotalFunction.StandardDeviation);

        if (stdDevPSubtotal)
            subtotals.Add(XlsxSubtotalFunction.PopulationStandardDeviation);

        if (varSubtotal)
            subtotals.Add(XlsxSubtotalFunction.Variance);

        if (varPSubtotal)
            subtotals.Add(XlsxSubtotalFunction.PopulationVariance);

        var xlReference = new XlsxPivotReference
        {
            Field = field,
            Selected = selected,
            ByPosition = byPosition,
            Relative = relative,
            Subtotals = subtotals,
        };

        // Add indexes after the reference is initialized, so it can check values by cacheIndex/byPosition.
        foreach (var fieldItem in Children(reference, "x"))
        {
            var fieldItemValue = (uint?)fieldItem.Attribute("v") ?? throw PartStructureException.MissingAttribute();
            xlReference.AddFieldItem(fieldItemValue);
        }

        return xlReference;
    }

    private static void LoadPivotTableStyle(XElement? pivotTableStyle, XlsxPivotTable xlPivotTable)
    {
        if (pivotTableStyle is not null)
        {
            xlPivotTable.Theme = Enum.TryParse<XlsxPivotTableTheme>((string?)pivotTableStyle.Attribute("name"), out var xlPivotTableTheme)
                ? xlPivotTableTheme
                : XlsxPivotTableTheme.None;
            xlPivotTable.ShowRowHeaders = (bool?)pivotTableStyle.Attribute("showRowHeaders") ?? false;
            xlPivotTable.ShowColumnHeaders = (bool?)pivotTableStyle.Attribute("showColHeaders") ?? false;
            xlPivotTable.ShowRowStripes = (bool?)pivotTableStyle.Attribute("showRowStripes") ?? false;
            xlPivotTable.ShowColumnStripes = (bool?)pivotTableStyle.Attribute("showColStripes") ?? false;
            xlPivotTable.ShowLastColumn = xlPivotTable.ShowColumnStripes;
        }
    }

    private static void LoadExtensionList(XElement pivotTable, XlsxPivotTable xlPivotTable)
    {
        var ptExt2010 = Extension(pivotTable, "pivotTableDefinition");
        if (ptExt2010 is not null)
        {
            xlPivotTable.EnableCellEditing = (bool?)ptExt2010.Attribute("enableEdit") ?? false;
            var hideValuesRow = (bool?)ptExt2010.Attribute("hideValuesRow") ?? false;
            xlPivotTable.ShowValuesRow = !hideValuesRow;
        }
    }

    private static XlsxPivotSummary? ParseSummary(string? value) => value switch
    {
        null => null,
        "sum" => XlsxPivotSummary.Sum,
        "count" => XlsxPivotSummary.Count,
        "average" => XlsxPivotSummary.Average,
        "min" => XlsxPivotSummary.Minimum,
        "max" => XlsxPivotSummary.Maximum,
        "product" => XlsxPivotSummary.Product,
        "countNums" => XlsxPivotSummary.CountNumbers,
        "stdDev" => XlsxPivotSummary.StandardDeviation,
        "stdDevp" => XlsxPivotSummary.PopulationStandardDeviation,
        "var" => XlsxPivotSummary.Variance,
        "varp" => XlsxPivotSummary.PopulationVariance,
        _ => throw PartStructureException.InvalidAttributeValue(value)
    };

    private static XlsxPivotCalculation? ParseCalculation(string? value) => value switch
    {
        null => null,
        "normal" => XlsxPivotCalculation.Normal,
        "difference" => XlsxPivotCalculation.DifferenceFrom,
        "percent" => XlsxPivotCalculation.PercentageOf,
        "percentDiff" => XlsxPivotCalculation.PercentageDifferenceFrom,
        "runTotal" => XlsxPivotCalculation.RunningTotal,
        "percentOfRow" => XlsxPivotCalculation.PercentageOfRow,
        "percentOfCol" => XlsxPivotCalculation.PercentageOfColumn,
        "percentOfTotal" => XlsxPivotCalculation.PercentageOfTotal,
        "index" => XlsxPivotCalculation.Index,
        _ => throw PartStructureException.InvalidAttributeValue(value)
    };

    private static XlsxPivotFormatAction? ParseFormatAction(string? value) => value switch
    {
        null => null,
        "blank" => XlsxPivotFormatAction.Blank,
        "formatting" => XlsxPivotFormatAction.Formatting,
        _ => throw PartStructureException.InvalidAttributeValue(value)
    };

    private static XlsxPivotCfScope? ParseCfScope(string? value) => value switch
    {
        null => null,
        "selection" => XlsxPivotCfScope.SelectedCells,
        "data" => XlsxPivotCfScope.DataFields,
        "field" => XlsxPivotCfScope.FieldIntersections,
        _ => throw PartStructureException.InvalidAttributeValue(value)
    };

    private static XlsxPivotCfRuleType? ParseCfRuleType(string? value) => value switch
    {
        null => null,
        "all" => XlsxPivotCfRuleType.All,
        "column" => XlsxPivotCfRuleType.Column,
        "none" => XlsxPivotCfRuleType.None,
        "row" => XlsxPivotCfRuleType.Row,
        _ => throw PartStructureException.InvalidAttributeValue(value)
    };

    private static XlsxPivotSortType? ParseSortType(string? value) => value switch
    {
        null => null,
        "manual" => XlsxPivotSortType.Default,
        "ascending" => XlsxPivotSortType.Ascending,
        "descending" => XlsxPivotSortType.Descending,
        _ => throw PartStructureException.InvalidAttributeValue(value)
    };

    private static XlsxPivotItemType? ParseItemType(string? value) => value switch
    {
        null => null,
        "avg" => XlsxPivotItemType.Avg,
        "blank" => XlsxPivotItemType.Blank,
        "count" => XlsxPivotItemType.Count,
        "countA" => XlsxPivotItemType.CountA,
        "data" => XlsxPivotItemType.Data,
        "default" => XlsxPivotItemType.Default,
        "grand" => XlsxPivotItemType.Grand,
        "max" => XlsxPivotItemType.Max,
        "min" => XlsxPivotItemType.Min,
        "product" => XlsxPivotItemType.Product,
        "stdDev" => XlsxPivotItemType.StdDev,
        "stdDevP" => XlsxPivotItemType.StdDevP,
        "sum" => XlsxPivotItemType.Sum,
        "var" => XlsxPivotItemType.Var,
        "varP" => XlsxPivotItemType.VarP,
        _ => throw PartStructureException.InvalidAttributeValue(value)
    };

    private static XlsxPivotAreaType? ParseAreaType(string? value) => value switch
    {
        null => null,
        "none" => XlsxPivotAreaType.None,
        "normal" => XlsxPivotAreaType.Normal,
        "data" => XlsxPivotAreaType.Data,
        "all" => XlsxPivotAreaType.All,
        "origin" => XlsxPivotAreaType.Origin,
        "button" => XlsxPivotAreaType.Button,
        "topRight" => XlsxPivotAreaType.TopRight,
        "topEnd" => XlsxPivotAreaType.TopEnd,
        _ => throw PartStructureException.InvalidAttributeValue(value)
    };

    private static XlsxPivotAxis? ParseAxis(string? value) => value switch
    {
        null => null,
        "axisRow" => XlsxPivotAxis.AxisRow,
        "axisCol" => XlsxPivotAxis.AxisCol,
        "axisPage" => XlsxPivotAxis.AxisPage,
        "axisValues" => XlsxPivotAxis.AxisValues,
        _ => throw PartStructureException.InvalidAttributeValue(value)
    };
}
