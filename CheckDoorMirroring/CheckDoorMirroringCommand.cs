using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using static System.Net.Mime.MediaTypeNames;
using System.Xml.Linq;

namespace CheckDoorMirroring
{
    public class ParameterChoice
    {
        public string Name { get; set; }
        public ElementId Id { get; set; }
    }

    [Transaction(TransactionMode.Manual)]
    public class CheckDoorMirroringCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            Document doc = uiApp.ActiveUIDocument.Document;

            string selectedParamName = null;
            ElementId selectedParamId = null;

            // --- ШАГ 1: ПОКАЗЫВАЕМ ПЕРВОЕ ОКНО ВЫБОРА ПУТИ ---
            TaskDialog mainDialog = new TaskDialog("Выбор способа записи");
            mainDialog.MainInstruction = "Куда записать информацию о зеркальности?";
            mainDialog.MainContent = "Вы можете использовать стандартный системный параметр 'Комментарии' или выбрать другой текстовый параметр из списка.";

            mainDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Использовать параметр 'Комментарии'");
            mainDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "Выбрать другой параметр вручную...");
            mainDialog.CommonButtons = TaskDialogCommonButtons.Cancel;
            mainDialog.DefaultButton = TaskDialogResult.CommandLink1;

            TaskDialogResult mainResult = mainDialog.Show();

            // --- ШАГ 2: ОБРАБАТЫВАЕМ ВЫБОР ПОЛЬЗОВАТЕЛЯ ---
            switch (mainResult)
            {
                case TaskDialogResult.CommandLink1:
                    selectedParamName = "Комментарии";
                    selectedParamId = new ElementId(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);
                    break;

                case TaskDialogResult.CommandLink2:
                    List<ParameterChoice> availableParameters = GetAvailableParameters(doc);
                    if (!availableParameters.Any())
                    {
                        TaskDialog.Show("Ошибка", "В проекте не найдено ни одного подходящего текстового параметра.");
                        return Result.Failed;
                    }

                    TaskDialog selectionDialog = new TaskDialog("Выбор параметра");
                    selectionDialog.MainInstruction = "Выберите параметр для записи значения";
                    selectionDialog.MainContent = "Пожалуйста, выберите один из найденных текстовых параметров.";

                    int itemsToShow = Math.Min(availableParameters.Count, 100);
                    if (availableParameters.Count > 100)
                    {
                        selectionDialog.MainContent += $"\n(Показаны первые 100 из {availableParameters.Count} найденных параметров)";
                    }

                    for (int i = 0; i < itemsToShow; i++)
                    {
                        selectionDialog.AddCommandLink((TaskDialogCommandLinkId)(i + 1), availableParameters[i].Name);
                    }

                    selectionDialog.CommonButtons = TaskDialogCommonButtons.Cancel;
                    selectionDialog.DefaultButton = TaskDialogResult.Cancel;

                    TaskDialogResult selectionResult = selectionDialog.Show();

                    if (selectionResult == TaskDialogResult.Cancel) return Result.Cancelled;

                    int selectedIndex = (int)selectionResult - 1;
                    if (selectedIndex < 0 || selectedIndex >= itemsToShow) return Result.Failed;

                    ParameterChoice selectedChoice = availableParameters[selectedIndex];
                    selectedParamName = selectedChoice.Name;
                    selectedParamId = selectedChoice.Id;
                    break;

                case TaskDialogResult.Cancel:
                default:
                    return Result.Cancelled;
            }

            if (string.IsNullOrEmpty(selectedParamName) || selectedParamId == null || selectedParamId == ElementId.InvalidElementId)
            {
                TaskDialog.Show("Ошибка", "Не удалось определить параметр для записи.");
                return Result.Failed;
            }

            // --- ШАГ 3: ОСНОВНАЯ ЛОГИКА С ВЫБРАННЫМ ПАРАМЕТРОМ ---
            try
            {
                using (Transaction trans = new Transaction(doc, "Проверка зеркальности"))
                {
                    trans.Start();

                    if (mainResult == TaskDialogResult.CommandLink1)
                    {
                        ProcessElementsWithBuiltIn(doc, BuiltInCategory.OST_Doors, BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);
                        ProcessElementsWithBuiltIn(doc, BuiltInCategory.OST_Windows, BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);
                    }
                    else
                    {
                        ProcessElements(doc, BuiltInCategory.OST_Doors, selectedParamName);
                        ProcessElements(doc, BuiltInCategory.OST_Windows, selectedParamName);
                    }

                    string doorFilterName = $"Проверка зеркальности дверей ({selectedParamName})";
                    string windowFilterName = $"Проверка зеркальности окон ({selectedParamName})";

                    ElementId doorFilterId = CreateFilter(doc, doorFilterName, BuiltInCategory.OST_Doors, selectedParamId);
                    ElementId windowFilterId = CreateFilter(doc, windowFilterName, BuiltInCategory.OST_Windows, selectedParamId);

                    ApplyTemporaryViewSettings(uiApp.ActiveUIDocument.ActiveView, doorFilterId, windowFilterId);

                    trans.Commit();
                }
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Ошибка", $"Произошла ошибка: {ex.Message}\n{ex.StackTrace}");
                return Result.Failed;
            }
        }

        #region Вспомогательные методы

        private List<ParameterChoice> GetAvailableParameters(Document doc)
        {
            var choices = new List<ParameterChoice>();
            var processedParamIds = new HashSet<ElementId>();

            Category doorCat = doc.Settings.Categories.get_Item(BuiltInCategory.OST_Doors);
            Category windowCat = doc.Settings.Categories.get_Item(BuiltInCategory.OST_Windows);

            if (doorCat == null || windowCat == null) return choices;

            BindingMap bindingMap = doc.ParameterBindings;
            var it = bindingMap.ForwardIterator();
            it.Reset();

            while (it.MoveNext())
            {
                if (!(it.Current is InstanceBinding binding)) continue;

                Definition def = it.Key;
                if (def.GetDataType() != SpecTypeId.String.Text) continue;

                CategorySet categories = binding.Categories;
                if (categories.Contains(doorCat) && categories.Contains(windowCat))
                {
                    ParameterElement paramElem = FindParameterElement(doc, def);
                    if (paramElem != null && processedParamIds.Add(paramElem.Id))
                    {
                        choices.Add(new ParameterChoice { Name = def.Name, Id = paramElem.Id });
                    }
                }
            }

            if (!choices.Any(p => p.Id.IntegerValue == (int)BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS))
            {
                ElementId commentsParamId = new ElementId(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);
                ParameterElement commentsParamElement = doc.GetElement(commentsParamId) as ParameterElement;

                if (commentsParamElement != null)
                {
                    if (processedParamIds.Add(commentsParamId))
                    {
                        choices.Add(new ParameterChoice { Name = "Комментарии", Id = commentsParamId });
                    }
                }
            }
            return choices.OrderBy(c => c.Name).ToList();
        }

        private ParameterElement FindParameterElement(Document doc, Definition def)
        {
            if (def is ExternalDefinition extDef)
            {
                return SharedParameterElement.Lookup(doc, extDef.GUID);
            }
            else
            {
                return new FilteredElementCollector(doc)
                    .OfClass(typeof(ParameterElement))
                    .Cast<ParameterElement>()
                    .FirstOrDefault(p => p.GetDefinition()?.Name == def.Name);
            }
        }

        private void ProcessElements(Document doc, BuiltInCategory category, string paramName)
        {
            var collector = new FilteredElementCollector(doc)
                .OfCategory(category)
                .WhereElementIsNotElementType();

            foreach (var fi in collector.Cast<FamilyInstance>())
            {
                if (fi == null) continue;
                string status = fi.Mirrored ? "Отзеркалено" : "Верное отображение";
                Parameter param = fi.LookupParameter(paramName);
                if (param != null && !param.IsReadOnly)
                {
                    param.Set(status);
                }
            }
        }

        private void ProcessElementsWithBuiltIn(Document doc, BuiltInCategory category, BuiltInParameter builtInParam)
        {
            var collector = new FilteredElementCollector(doc)
                .OfCategory(category)
                .WhereElementIsNotElementType();

            foreach (var fi in collector.Cast<FamilyInstance>())
            {
                if (fi == null) continue;
                string status = fi.Mirrored ? "Отзеркалено" : "Верное отображение";
                Parameter param = fi.get_Parameter(builtInParam);
                if (param != null && !param.IsReadOnly)
                {
                    param.Set(status);
                }
            }
        }

        private ElementId CreateFilter(Document doc, string filterName, BuiltInCategory category, ElementId paramId)
        {
            var existingFilter = new FilteredElementCollector(doc)
                .OfClass(typeof(ParameterFilterElement))
                .Cast<ParameterFilterElement>()
                .FirstOrDefault(f => f.Name.Equals(filterName, StringComparison.OrdinalIgnoreCase));

            if (existingFilter != null) return existingFilter.Id;

            var categoryList = new List<ElementId> { doc.Settings.Categories.get_Item(category).Id };
            FilterRule rule = ParameterFilterRuleFactory.CreateEqualsRule(paramId, "Отзеркалено", false);

            ParameterFilterElement newFilter = ParameterFilterElement.Create(doc, filterName, categoryList);
            newFilter.SetElementFilter(new ElementParameterFilter(rule));

            return newFilter.Id;
        }

        private void ApplyTemporaryViewSettings(View view, ElementId doorFilterId, ElementId windowFilterId)
        {
            if (view == null || !view.CanEnableTemporaryViewPropertiesMode()) return;
            if (doorFilterId == ElementId.InvalidElementId && windowFilterId == ElementId.InvalidElementId) return;

            try
            {
                view.EnableTemporaryViewPropertiesMode(view.Id);
                ElementId solidPatternId = GetSolidFillPatternId(view.Document);

                var settings = new OverrideGraphicSettings();
                settings.SetProjectionLineColor(new Color(255, 0, 0));
                settings.SetCutLineColor(new Color(255, 0, 0));

                if (solidPatternId != null && solidPatternId != ElementId.InvalidElementId)
                {
                    settings.SetSurfaceForegroundPatternColor(new Color(255, 0, 0));
                    settings.SetSurfaceForegroundPatternId(solidPatternId);
                    settings.SetSurfaceForegroundPatternVisible(true);
                    settings.SetCutForegroundPatternColor(new Color(255, 0, 0));
                    settings.SetCutForegroundPatternId(solidPatternId);
                    settings.SetCutForegroundPatternVisible(true);
                }

                if (doorFilterId != ElementId.InvalidElementId) view.SetFilterOverrides(doorFilterId, settings);
                if (windowFilterId != ElementId.InvalidElementId) view.SetFilterOverrides(windowFilterId, settings);

                TaskDialog.Show("Успех", "Проверка зеркальности завершена. Временные переопределения графики применены для отзеркаленных элементов.");
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Ошибка применения вида", $"Не удалось применить временные переопределения: {ex.Message}");
                if (view.IsTemporaryViewPropertiesModeEnabled())
                {
                    view.DisableTemporaryViewMode(TemporaryViewMode.TemporaryViewProperties);
                }
            }
        }

        private ElementId GetSolidFillPatternId(Document doc)
        {
            return new FilteredElementCollector(doc)
                .OfClass(typeof(FillPatternElement))
                .Cast<FillPatternElement>()
                .FirstOrDefault(x => x.GetFillPattern().IsSolidFill)?.Id;
        }

        #endregion
    }
}
