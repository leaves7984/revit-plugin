using System.IO;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Net;

using RvtApplication = Autodesk.Revit.ApplicationServices.Application;
using RvtDocument = Autodesk.Revit.DB.Document;
using Newtonsoft.Json;

namespace RevitPlugin02
{
    [TransactionAttribute(TransactionMode.Manual)]
    class GetNetWallArea : IExternalCommand
    {
        const string _export_folder_name = "C:\\yuan20\\tmp\\schedule_data";
        const string root_api = "https://bimport.buildsim.io/backend/";
        private static ExternalCommandData _cachedCmdData;
        private static List<Product> list;
        public class Product
        {
            public string index { get; set; }
            public string count { get; set; }
            public Product(string index, string count)
            {
                this.index = index;
                this.count = count;
            }

        }

        public static UIApplication CachedUiApp
        {
            get
            {
                return _cachedCmdData.Application;
            }
        }

        public static RvtApplication CachedApp
        {
            get
            {
                return CachedUiApp.Application;
            }
        }

        public static RvtDocument CachedDoc
        {
            get
            {
                return CachedUiApp.ActiveUIDocument.Document;
            }
        }
        /// Add fields to view schedule.
        /// 

        /// List of view schedule.
        public void AddFieldToSchedule(IList schedules)
        {
            IList<SchedulableField> schedulableFields = null;

            foreach (ViewSchedule vs in schedules)
            {
                //Get all schedulable fields from view schedule definition. 
                schedulableFields = vs.Definition.GetSchedulableFields();

                foreach (SchedulableField sf in schedulableFields)
                {
                    bool fieldAlreadyAdded = false; //Get all schedule field ids IList 

                    List<ScheduleFieldId> ids = vs.Definition.GetFieldOrder().ToList();

                    foreach (ScheduleFieldId id in ids)
                    {
                        //If the GetSchedulableField() method of gotten schedule field returns same schedulable field,
                        // it means the field is already added to the view schedule. 
                        if (vs.Definition.GetField(id).GetSchedulableField() == sf)
                        {
                            fieldAlreadyAdded = true;
                            break;
                        }
                    } //If schedulable field doesn't exist in view schedule, add it. 

                    if (fieldAlreadyAdded == false)
                    {
                        vs.Definition.AddField(sf);
                    }
                }
            }
        }
        public void createMaterialTakeOffSchedule( String name, ElementId elementId)
        {

            UIDocument uidoc = _cachedCmdData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            ViewSchedule schedule = new FilteredElementCollector(CachedDoc).OfClass(typeof(ViewSchedule)).Where(x => x.Name == name + " Schedule").FirstOrDefault() as ViewSchedule;
            if (schedule == null)
            {
                Transaction tSchedule = new Transaction(CachedDoc, "Create Schedule");
                tSchedule.Start();

                //Create an empty view schedule for doors.
                schedule = ViewSchedule.CreateMaterialTakeoff(CachedDoc, elementId);
                schedule.Name = name + " Schedule";

                ElementId volumeId = new ElementId(BuiltInParameter.MATERIAL_VOLUME);
                ElementId areaId = new ElementId(BuiltInParameter.MATERIAL_AREA);
                //Iterate all the schedulable fields gotten from the doors view schedule.
                foreach (SchedulableField schedulableField in schedule.Definition.GetSchedulableFields())
                {
                    //See if the FieldType is ScheduleFieldType.Instance.
                    if (schedulableField.ParameterId == volumeId || schedulableField.ParameterId == areaId || schedulableField.GetName(doc).Equals("Material: Keynote"))
                    {
                        //Get ParameterId of SchedulableField.
                        ElementId parameterId = schedulableField.ParameterId;

                        //Add a new schedule field to the view schedule by using the SchedulableField as argument of AddField method of Autodesk.Revit.DB.ScheduleDefinition class.
                        ScheduleField field = schedule.Definition.AddField(schedulableField);

                        //See if the parameterId is a BuiltInParameter.
                        if (Enum.IsDefined(typeof(BuiltInParameter), parameterId.IntegerValue))
                        {
                            BuiltInParameter bip = (BuiltInParameter)parameterId.IntegerValue;
                            //Get the StorageType of BuiltInParameter.
                            Autodesk.Revit.DB.StorageType st = CachedDoc.get_TypeOfStorage(bip);
                            //if StorageType is String or ElementId, set GridColumnWidth of schedule field to three times of current GridColumnWidth.
                            //And set HorizontalAlignment property to left.
                            if (st == Autodesk.Revit.DB.StorageType.String || st == Autodesk.Revit.DB.StorageType.ElementId)
                            {
                                field.GridColumnWidth = 3 * field.GridColumnWidth;
                                field.HorizontalAlignment = ScheduleHorizontalAlignment.Left;
                            }
                            //For other StorageTypes, set HorizontalAlignment property to center.
                            else
                            {
                                field.HorizontalAlignment = ScheduleHorizontalAlignment.Center;
                            }
                        }

                        if (schedulableField.GetName(doc).Equals("Material: Keynote"))
                        {
                            ScheduleSortGroupField sortGroupField = new ScheduleSortGroupField(field.FieldId);
                            schedule.Definition.AddSortGroupField(sortGroupField);
                            schedule.Definition.IsItemized = false;
                        }

                        if (field.ParameterId == volumeId)
                        {
                            FormatOptions formatOpt = new FormatOptions(DisplayUnitType.DUT_CUBIC_FEET, UnitSymbolType.UST_CF, 0.01);
                            formatOpt.UseDefault = false;
                            field.SetFormatOptions(formatOpt);
                            field.DisplayType = ScheduleFieldDisplayType.Totals;
                        }

                        if (field.ParameterId == areaId)
                        {
                            FormatOptions formatOpt = new FormatOptions(DisplayUnitType.DUT_SQUARE_FEET, UnitSymbolType.UST_SF, 0.01);
                            formatOpt.UseDefault = false;
                            field.SetFormatOptions(formatOpt);
                            field.DisplayType = ScheduleFieldDisplayType.Totals;
                        }

                    }
                }


                tSchedule.Commit();
                tSchedule.Dispose();

            }
            else
            {
                schedule.RefreshData();
            }

            ViewScheduleExportOptions opt = new ViewScheduleExportOptions();
            opt.FieldDelimiter = ",";
            opt.Title = false;

            string path = _export_folder_name;


            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);


            string file = System.IO.Path.GetFileNameWithoutExtension(name) + ".csv";

            schedule.Export(path, file, opt);
        }
        public void createLengthSchedule(String name, ElementId elementId)
        {
            ViewSchedule schedule = new FilteredElementCollector(CachedDoc).OfClass(typeof(ViewSchedule)).Where(x => x.Name == name + " Schedule").FirstOrDefault() as ViewSchedule;
            if (schedule == null)
            {
                Transaction tSchedule = new Transaction(CachedDoc, "Create Schedule");
                tSchedule.Start();

                //Create an empty view schedule for doors.
                schedule = ViewSchedule.CreateSchedule(CachedDoc, elementId, ElementId.InvalidElementId);
                schedule.Name = name + " Schedule";

                ElementId keynoteId = new ElementId(BuiltInParameter.KEYNOTE_PARAM);
                ElementId lengthId = new ElementId(BuiltInParameter.CURVE_ELEM_LENGTH);
                //Iterate all the schedulable fields gotten from the doors view schedule.
                foreach (SchedulableField schedulableField in schedule.Definition.GetSchedulableFields())
                {
                    //See if the FieldType is ScheduleFieldType.Instance.
                    if (schedulableField.ParameterId == lengthId || schedulableField.ParameterId == keynoteId)
                    {
                        //Get ParameterId of SchedulableField.
                        ElementId parameterId = schedulableField.ParameterId;

                        //Add a new schedule field to the view schedule by using the SchedulableField as argument of AddField method of Autodesk.Revit.DB.ScheduleDefinition class.
                        ScheduleField field = schedule.Definition.AddField(schedulableField);

                        //See if the parameterId is a BuiltInParameter.
                        if (Enum.IsDefined(typeof(BuiltInParameter), parameterId.IntegerValue))
                        {
                            BuiltInParameter bip = (BuiltInParameter)parameterId.IntegerValue;
                            //Get the StorageType of BuiltInParameter.
                            Autodesk.Revit.DB.StorageType st = CachedDoc.get_TypeOfStorage(bip);
                            //if StorageType is String or ElementId, set GridColumnWidth of schedule field to three times of current GridColumnWidth.
                            //And set HorizontalAlignment property to left.
                            if (st == Autodesk.Revit.DB.StorageType.String || st == Autodesk.Revit.DB.StorageType.ElementId)
                            {
                                field.GridColumnWidth = 3 * field.GridColumnWidth;
                                field.HorizontalAlignment = ScheduleHorizontalAlignment.Left;
                            }
                            //For other StorageTypes, set HorizontalAlignment property to center.
                            else
                            {
                                field.HorizontalAlignment = ScheduleHorizontalAlignment.Center;
                            }
                        }

                        if (field.ParameterId == keynoteId)
                        {
                            ScheduleSortGroupField sortGroupField = new ScheduleSortGroupField(field.FieldId);
                            schedule.Definition.AddSortGroupField(sortGroupField);
                            schedule.Definition.IsItemized = false;
                        }
                        if (field.ParameterId == lengthId)
                        {

                            field.DisplayType = ScheduleFieldDisplayType.Totals;
                        }
                    }
                }


                tSchedule.Commit();
                tSchedule.Dispose();

            }
            else
            {
                schedule.RefreshData();
            }

            ViewScheduleExportOptions opt = new ViewScheduleExportOptions();
            opt.FieldDelimiter = ",";
            opt.Title = false;

            string path = _export_folder_name;


            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);


            string file = System.IO.Path.GetFileNameWithoutExtension(name) + ".csv";

            schedule.Export(path, file, opt);
        }
        public void createCountSchedule(String name, ElementId elementId)
        {
            ViewSchedule schedule = new FilteredElementCollector(CachedDoc).OfClass(typeof(ViewSchedule)).Where(x => x.Name == name + " Schedule").FirstOrDefault() as ViewSchedule;
            if (schedule == null)
            {
                Transaction tSchedule = new Transaction(CachedDoc, "Create Schedule");
                tSchedule.Start();

                //Create an empty view schedule for doors.
                //schedule = ViewSchedule.CreateSchedule(CachedDoc, new ElementId(BuiltInCategory.INVALID), ElementId.InvalidElementId);
                schedule = ViewSchedule.CreateSchedule(CachedDoc, elementId, ElementId.InvalidElementId);
                schedule.Name = name + " Schedule";

                ElementId keynoteId = new ElementId(BuiltInParameter.KEYNOTE_PARAM);
                //Iterate all the schedulable fields gotten from the doors view schedule.
                foreach (SchedulableField schedulableField in schedule.Definition.GetSchedulableFields())
                {
                    //See if the FieldType is ScheduleFieldType.Instance.
                    if (schedulableField.FieldType == ScheduleFieldType.Count || schedulableField.ParameterId == keynoteId)
                    {
                        //Get ParameterId of SchedulableField.
                        ElementId parameterId = schedulableField.ParameterId;

                        //Add a new schedule field to the view schedule by using the SchedulableField as argument of AddField method of Autodesk.Revit.DB.ScheduleDefinition class.
                        ScheduleField field = schedule.Definition.AddField(schedulableField);

                        //See if the parameterId is a BuiltInParameter.
                        if (Enum.IsDefined(typeof(BuiltInParameter), parameterId.IntegerValue))
                        {
                            BuiltInParameter bip = (BuiltInParameter)parameterId.IntegerValue;
                            //Get the StorageType of BuiltInParameter.
                            Autodesk.Revit.DB.StorageType st = CachedDoc.get_TypeOfStorage(bip);
                            //if StorageType is String or ElementId, set GridColumnWidth of schedule field to three times of current GridColumnWidth.
                            //And set HorizontalAlignment property to left.
                            if (st == Autodesk.Revit.DB.StorageType.String || st == Autodesk.Revit.DB.StorageType.ElementId)
                            {
                                field.GridColumnWidth = 3 * field.GridColumnWidth;
                                field.HorizontalAlignment = ScheduleHorizontalAlignment.Left;
                            }
                            //For other StorageTypes, set HorizontalAlignment property to center.
                            else
                            {
                                field.HorizontalAlignment = ScheduleHorizontalAlignment.Center;
                            }
                        }

                        if (field.ParameterId == keynoteId)
                        {
                            ScheduleSortGroupField sortGroupField = new ScheduleSortGroupField(field.FieldId);
                            schedule.Definition.AddSortGroupField(sortGroupField);
                            schedule.Definition.IsItemized = false;
                        }
                        if (schedulableField.FieldType == ScheduleFieldType.Count)
                        {

                        }

                    }
                }


                tSchedule.Commit();
                tSchedule.Dispose();

            }
            else
            {
                schedule.RefreshData();
            }

            ViewScheduleExportOptions opt = new ViewScheduleExportOptions();
            opt.FieldDelimiter = ",";
            opt.Title = false;

            string path = _export_folder_name;


            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);


            string file = System.IO.Path.GetFileNameWithoutExtension(name) + ".csv";

            schedule.Export(path, file, opt);
        }


        public static void readExcelFromFolder()
        {
            String[] dir = Directory.GetFiles(_export_folder_name, "*.csv");
            System.Diagnostics.Debug.WriteLine("path:::");
            foreach (String path in dir)
            {

                // length schedule
                if (path.Contains("-2"))
                {

                }
                // material takeoff
                else if (path.Contains("-3"))
                {
                    readMaterialTakeOff(path);
                }
                // count schedule
                else
                {
                    readCountSchedule(path);
                }

            }


        }
        public static void readMaterialTakeOff(String path)
        {
            FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.None);
            StreamReader sr = new StreamReader(fs);
            string str = "";
            string head = sr.ReadLine();
            String[] lines = null;
            if (head != null)
            {
                sr.ReadLine();
                while ((str = sr.ReadLine()) != null)
                {
                    str = str.Replace("\"", "");
                    lines = str.Split(',');
                    if (lines.Length == 3)
                    {
                        String area = lines[0];
                        String volume = lines[1];
                        String keynote = lines[2];
                        if (keynote.Length > 2)
                        {
                            Product p = new Product(keynote, area);
                            list.Add(p);
                            System.Diagnostics.Debug.WriteLine(keynote + ":" + area);
                        }

                    }


                }
            }

            sr.Close();
            fs.Close();
        }
        public static void readCountSchedule(String path)
        {
            FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.None);
            StreamReader sr = new StreamReader(fs);
            string str = "";
            string head = sr.ReadLine();
            String[] lines = null;
            if (head != null)
            {
                sr.ReadLine();
                while ((str = sr.ReadLine()) != null)
                {
                    str = str.Replace("\"", "");
                    lines = str.Split(',');
                    if (lines.Length == 2)
                    {
                        String count = lines[0];
                        String keynote = lines[1];
                        if (keynote.Length > 2)
                        {
                            Product p = new Product(keynote, count);
                            list.Add(p);
                            System.Diagnostics.Debug.WriteLine(keynote + ":" + count);
                        }

                    }

                }
            }

            sr.Close();
            fs.Close();

        }
        
        public void getNetWallArea()
        {
            // Get UIDocument
            UIDocument uidoc = _cachedCmdData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
           
            var walls = new FilteredElementCollector(doc).OfClass(typeof(Wall)).Cast<Wall>();
            
            foreach (Wall w in walls)
            {
                Reference sideFaceRef = HostObjectUtils.GetSideFaces(w, ShellLayerType.Exterior).First();
                Face netFace = w.GetGeometryObjectFromReference(sideFaceRef) as Face;

                double netArea = netFace.Area;
                double grossArea;
                
                using (Transaction t = new Transaction(doc, "delete inserts"))
                {
                    t.Start();
                    doc = uidoc.Document;
                    //var instances = new FilteredElementCollector(doc).OfClass(typeof(FamilyInstance)).Cast<FamilyInstance>().Where(q => q.Host != null && q.Host.Id == w.Id);
                    var ids = w.FindInserts(true, true, true, true);
                    foreach (ElementId id in ids)
                    {
                        var el = doc.GetElement(id);
                        if (el is FamilyInstance) {
                            var fi = el as FamilyInstance;
                            if (fi != null) {
                                doc.Delete(fi.Id);
                            }
                        }
                        
                    }
                    
                    //doc.Regenerate();
                    Face grossFace = w.GetGeometryObjectFromReference(sideFaceRef) as Face;
                    grossArea = grossFace.Area;
            
                    t.Commit();
                }
                TaskDialog.Show("area", "net = " + netArea + "\nGross = " + grossArea);
            }
            

        }

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {

            _cachedCmdData = commandData;
            try
            {
                //TODO: add your code below.
          
                getNetWallArea();
                createMaterialTakeOffSchedule("Walls-3", new ElementId(BuiltInCategory.OST_Walls));

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.ToString();
                TaskDialog.Show("Cost", string.Format("total ${0}", "error"));
                return Result.Failed;
            }
        }
    }
}
