using System;
using System.Collections.Generic;
using System.Text;
using System.Reflection;
using System.ComponentModel;
using System.Collections;

namespace Taramon.Exceller
{
    public class ExcelMapper
    {
        public void Read(object destination, string excelFileName)
        {
            if (destination == null)
                throw new ArgumentNullException("destination");

            using (ExcelManager em = new ExcelManager())
            {
                em.Open(excelFileName);

                HandleCommonClassAttributes(em, destination);

                // Check attributes of Properties in class
                Type objType = destination.GetType();
                foreach (PropertyInfo propInfo in objType.GetProperties())
                {
                    object[] attributes = propInfo.GetCustomAttributes(true);
                    if (attributes == null)
                        continue;

                    foreach (Attribute att in attributes)
                    {
                        HandleCommonPropertiesAttributes(em, destination, att, propInfo);

                        // Read Specific Attributes:

                        // [FromCell]
                        FromCell fromCell = att as FromCell;
                        if (fromCell != null)
                        {
                            object value = em.GetValue(fromCell.CellAddress, fromCell.Category);
                            propInfo.SetValue(destination, value, null);
                        }

                        // [FromRange]
                        FromRange fromRange = att as FromRange;
                        if (fromRange != null)
                        {
                            ArrayList values = em.GetRangeValues(fromRange.StartCellAddress, fromRange.EndCellAddress, fromRange.Category);
                            propInfo.SetValue(destination, values, null);
                        }
                    }
                }
            }
        }

        public void Write(object target, string excelFileName)
        {
            if (target == null)
                throw new ArgumentNullException("target");

            using (ExcelManager em = new ExcelManager())
            {
                em.Create(excelFileName);
                Write(em, target);
                em.Save();
            }
        }

        public void Write(object target, string excelFileName, string templateFileName)
        {
            if (target == null)
                throw new ArgumentNullException("target");

            using (ExcelManager em = new ExcelManager())
            {
                em.Open(templateFileName);
                Write(em, target);
                em.SaveAs(excelFileName);
            }
        }

        private void Write(ExcelManager em, object target)
        {
            HandleCommonClassAttributes(em, target);
            // Check attributes of Properties in class
            Type objType = target.GetType();
            foreach (PropertyInfo propInfo in objType.GetProperties())
            {
                object[] attributes = propInfo.GetCustomAttributes(true);
                if (attributes == null)
                    continue;

                foreach (Attribute att in attributes)
                {
                    HandleCommonPropertiesAttributes(em, target, att, propInfo);

                    // Write Specific Attributes:

                    // [ToCell]
                    ToCell toCell = att as ToCell;
                    if (toCell != null)
                    {
                        em.SetValue(toCell.CellAddress, propInfo.GetValue(target, null));
                        continue; // Do not check other types of Attribute (speed-optimization)
                    }

                    // [ToRange]
                    ToRange toRange = att as ToRange;
                    if (toRange != null)
                    {
                        if (propInfo.PropertyType == typeof(ArrayList))
                        {
                            ArrayList temporaryArray = propInfo.GetValue(target, null) as ArrayList;
                            em.SetRangeValues(toRange.StartCellAddress, toRange.EndCellAddress, temporaryArray);
                        }
                        else
                        {
                            // It's a non-array value should be repeated in all cells of the range
                            em.SetRangeValue(
                                toRange.StartCellAddress,
                                toRange.EndCellAddress,
                                propInfo.GetValue(target, null));
                        }
                    }
                }
            }
        }

        private void HandleCommonClassAttributes(ExcelManager em, object obj)
        {
            Type objType = obj.GetType();
            // Check class attributes
            foreach (Attribute att in objType.GetCustomAttributes(true))
            {
                // [DefaultSheet]
                DefaultSheet defaultSheet = att as DefaultSheet;
                if (defaultSheet != null)
                {
                    em.ActivateSheet(defaultSheet.SheetName);
                }
            }
        }

        private void HandleCommonPropertiesAttributes(ExcelManager em, object obj, Attribute att, PropertyInfo propInfo)
        {
            UseSheet useSheet = att as UseSheet;
            if (useSheet != null)
            {
                if (!String.IsNullOrEmpty(useSheet.SheetName))
                {
                    em.ActivateSheet(useSheet.SheetName);
                }
                else
                {
                    string sheetName = Convert.ToString(propInfo.GetValue(obj, null));
                    em.ActivateSheet(sheetName);
                }
            }
        }
    }
}
