using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DAL_SQL_Server.Entities;
using System.IO;

namespace DAL_File_Factory.Factory.CREATE_ENTITY
{
    public class CreationFactory
    {
        public static List<NewColumns> Create(string FileNameIn, string DelimiterIn, bool blnFirstRowContainsColumnNames, bool blnUpdateDataTypesDependentOnDataInFile, bool blnAllowNullableDataTypes)
        {
            List<NewColumns> response = new List<NewColumns>();
            try
            {
                if (!blnFirstRowContainsColumnNames)
                    throw new ApplicationException("Excepted true for blnFirstRowContainsColumnNames (not implemented for files without column names for first row)");

                string[] lines = null;
                lines = File.ReadAllLines(FileNameIn);

                if (lines.Length == 0)
                    throw new ApplicationException(string.Format("No Data Error. File '{0}'", FileNameIn));

                string          line = null;                
                List<string>    data = null;
                string          text = null;
                
                line = lines[0];
                data = line.Split(DelimiterIn[0]).ToList();

                for (Int32 i = 0; i < data.Count; i++)
                {
                    string strColumnName = null;
                    strColumnName = data[i];

                    NewColumns obj = new NewColumns();
                    obj.COLUMN_NAME         = strColumnName;
                    obj.ORDINAL_POSITION    = i;
                    obj.DATA_TYPE           = "varchar";
                }

                if (!blnUpdateDataTypesDependentOnDataInFile)
                    return response;
                
                //update data type dependent on data                
                foreach (NewColumns obj in response)
                {
                    bool IsDateTime = true;
                    bool IsDecimal = true;
                    bool IsInt32 = true;
                    bool IsInt16 = true;
                    bool IsNullable = false;
                    Int32 linesChecked = 0;
                    for (Int32 i = 1; i < lines.Length; i++)
                    {
                        data = line.Split(DelimiterIn[0]).ToList();
                        text = data[obj.ORDINAL_POSITION.Value];
                        if (text.Length == 0)
                        {
                            IsNullable = true;
                            continue;
                        }
                        linesChecked++;
                        if (IsDateTime)
                        {
                            try
                            {
                                DateTime.Parse(text);
                            }
                            catch (Exception ex)
                            {
                                IsDateTime = false;
                            }
                        }

                        if (IsDecimal)
                        {
                            try
                            {
                                Decimal.Parse(text);
                            }
                            catch (Exception ex)
                            {
                                IsDecimal = false;
                            }
                        }
                        if (IsInt32)
                        {
                            try
                            {
                                Int32.Parse(text);
                            }
                            catch (Exception ex)
                            {
                                IsInt32 = false;
                            }
                        }
                        if (IsInt16)
                        {
                            try
                            {
                                Int16.Parse(text);
                            }
                            catch (Exception ex)
                            {
                                IsInt16 = false;
                            }
                        }
                    }
                    
                    if (linesChecked > 0)
                    {
                        if (IsDateTime)
                        {
                            obj.DATA_TYPE = "datetime";
                        }
                        else if (IsInt16)
                        {
                            obj.DATA_TYPE = "numeric";
                            obj.NUMERIC_PRECISION = 4;
                        }
                        else if (IsInt32)
                        {
                            obj.DATA_TYPE = "numeric";
                            obj.NUMERIC_PRECISION = 9;
                        }
                        else if (IsInt16)
                        {
                            obj.DATA_TYPE = "numeric";
                            obj.NUMERIC_PRECISION = 18;
                        }
                        else
                        {   // string
                            obj.DATA_TYPE = "varchar";                            
                        }

                        if (IsNullable && blnAllowNullableDataTypes)
                        {
                            obj.IS_NULLABLE = "YES";
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return response; 
        }
    }
}
