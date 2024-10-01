using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.IO;

namespace RevitLines
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class PatternToModelLine : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;
            message = string.Empty;

            // Verificar se a vista ativa é 2D
            if (!(uiDoc.ActiveView is ViewPlan) && !(uiDoc.ActiveView is ViewSection))
            {
                message = "Por favor, selecione um elemento em uma vista 2D.";
                return Result.Failed;
            }

            // Selecionar o objeto na vista 2D
            Reference pickedRef = uiDoc.Selection.PickObject(ObjectType.Element, "Selecione um elemento");
            Element element = doc.GetElement(pickedRef);

            if (element == null)
            {
                message = "Elemento não encontrado.";
                return Result.Failed;
            }

            // Pegar a posição do elemento
            XYZ position = GetElementPosition(element);
            if (position == null)
            {
                message = "Não foi possível obter a posição do elemento.";
                return Result.Failed;
            }

            // Criar um arquivo DWG temporário
            string tempDwgFilePath = Path.GetTempFileName().Replace(".tmp", ".dwg");

            using (Transaction transaction = new Transaction(doc, "Exportar e Importar DWG"))
            {
                transaction.Start();

                try
                {
                    // Criar uma vista temporária
                    View tempView = CreateTemporaryView(doc, uiDoc.ActiveView, element);

                    // Exportar o elemento para o arquivo temporário
                    ExportToTemporaryDWG(doc, tempView, element, tempDwgFilePath);

                    // Importar o DWG temporário de volta para o Revit
                    ImportDwgToRevit(doc, tempDwgFilePath, position);

                    // Excluir a vista temporária
                    doc.Delete(tempView.Id);
                }
                catch (Exception ex)
                {
                    message = $"Falha ao exportar ou importar o DWG: {ex.Message}";
                    transaction.RollBack(); // Reverter a transação em caso de falha
                    return Result.Failed;
                }
                finally
                {
                    // Excluir o arquivo temporário após a importação
                    if (File.Exists(tempDwgFilePath))
                    {
                        File.Delete(tempDwgFilePath);
                    }
                }

                transaction.Commit();
            }

            return Result.Succeeded;
        }

        private XYZ GetElementPosition(Element element)
        {
            // Verifica se o elemento tem uma localização direta (LocationPoint ou LocationCurve)
            if (element.Location != null)
            {
                LocationPoint locationPoint = element.Location as LocationPoint;
                if (locationPoint != null)
                {
                    return locationPoint.Point;
                }

                LocationCurve locationCurve = element.Location as LocationCurve;
                if (locationCurve != null)
                {
                    return locationCurve.Curve.GetEndPoint(0);
                }
            }

            // Tenta obter a posição a partir da geometria do forro (se for um forro)
            if (element.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Ceilings)
            {
                Options geomOptions = new Options { ComputeReferences = true };
                GeometryElement geomElement = element.get_Geometry(geomOptions);

                if (geomElement != null)
                {
                    foreach (GeometryObject geomObj in geomElement)
                    {
                        Solid solid = geomObj as Solid;
                        if (solid != null && solid.Faces.Size > 0)
                        {
                            BoundingBoxXYZ bbox = solid.GetBoundingBox();
                            if (bbox != null)
                            {
                                XYZ center = (bbox.Min + bbox.Max) / 2;
                                return center;
                            }
                        }
                    }
                }
            }

            // Se o elemento for uma família, tenta obter a localização da instância
            FamilyInstance familyInstance = element as FamilyInstance;
            if (familyInstance != null)
            {
                LocationPoint familyLocationPoint = familyInstance.Location as LocationPoint;
                if (familyLocationPoint != null)
                {
                    return familyLocationPoint.Point;
                }
            }

            return null;
        }

        private View CreateTemporaryView(Document doc, View originalView, Element element)
        {
            // Duplicar a vista ativa como uma nova vista temporária
            ViewDuplicateOption duplicateOption = ViewDuplicateOption.Duplicate;
            View tempView = doc.GetElement(originalView.Duplicate(duplicateOption)) as View;

            // Isolar o elemento na nova vista temporária
            tempView.IsolateElementTemporary(element.Id); // Passando o ElementId diretamente

            return tempView;
        }

        private void ExportToTemporaryDWG(Document doc, View view, Element element, string dwgFilePath)
        {
            DWGExportOptions options = new DWGExportOptions
            {
                FileVersion = ACADVersion.R2013,
                ExportOfSolids = SolidGeometry.ACIS // Exportar sólidos
            };

            ICollection<ElementId> ids = new List<ElementId> { element.Id };

            // Exportar o elemento selecionado como DWG a partir da vista temporária
            doc.Export(Path.GetDirectoryName(dwgFilePath), Path.GetFileNameWithoutExtension(dwgFilePath), new List<ElementId> { view.Id }, options);
        }

        private void ImportDwgToRevit(Document doc, string dwgFilePath, XYZ position)
        {
            DWGImportOptions importOptions = new DWGImportOptions
            {
                Placement = ImportPlacement.Origin,
                CustomScale = 1.0,
                Unit = ImportUnit.Default
            };

            bool importSuccess = doc.Import(dwgFilePath, importOptions, doc.ActiveView, out ElementId importedElementId);

            if (importSuccess && importedElementId != ElementId.InvalidElementId)
            {
                ImportInstance importInstance = doc.GetElement(importedElementId) as ImportInstance;
                if (importInstance != null)
                {
                    // Verificar se o elemento está fixado (pinned)
                    bool wasPinned = importInstance.Pinned;

                    if (wasPinned)
                    {
                        // Desfixar o elemento antes de movê-lo
                        importInstance.Pinned = false;
                    }

                    // Mover o DWG para a posição correta
                    ElementTransformUtils.MoveElement(doc, importInstance.Id, position);

                    if (wasPinned)
                    {
                        // Refixar o elemento após movê-lo
                        importInstance.Pinned = true;
                    }
                }
                else
                {
                    throw new InvalidOperationException("Falha ao importar o arquivo DWG.");
                }
            }
            else
            {
                throw new InvalidOperationException("Falha ao importar o arquivo DWG.");
            }
        }
    }
}
