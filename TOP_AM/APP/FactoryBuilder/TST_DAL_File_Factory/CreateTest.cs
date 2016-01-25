using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using DAL_SQL_Server.Entities.INFORMATION_SCHEMA;
using DAL_SQL_Server.Entities;


namespace TST_DAL_File_Factory
{
    [TestClass]
    public class CreateTest
    {
        #region Attributes
        public const string cFileName = @"C:\Users\David\Desktop\ForWork\OMNI_Trades.txt";

        public const string ProjectPath = @"C:\Users\David\Desktop\ForWork\DAL_TOP_AM\DAL_TOP_AM";
        public const string NameSpace   = "DAL_TOP_AM";
        public const string EntityName  = "Trade_OMNI";
        public const string FactoryName = "OMNI";
        public const string TableName   = "Trade_OMNI";

        public static List<COLUMNS> listColumns = null;
        public static List<NewColumns> listMyColumns = null;


        #endregion
        #region Tests


        [ClassInitialize()]
        public static void MyClassInitialize(TestContext testContext)
        {
            DAL_SQL_Server.Factory.CREATE_ENTITY.CreationFactory.cDefaultCreateAndInsert = true;
            DAL_SQL_Server.Factory.CREATE_ENTITY.CreationFactory.cIsPartialSelectFactory = true;
            DAL_SQL_Server.Factory.CREATE_ENTITY.CreationFactory.cIsPartialEntity = true;
            DAL_SQL_Server.Factory.CREATE_ENTITY.CreationFactory.cIsPartialUpdateFactory = true;
            listColumns = DAL_SQL_Server.Factory.INFORMATION_SCHEMA.COLUMNS.SelectFactory.SelectByTABLE_NAMEorderByORDINAL_POSITION(TableName);
            listMyColumns = listColumns.ConvertBaseType();
        }

        [TestMethod]
        public void A_Create_FilterColumns__Operators_Table()
        {
            // not implemented
        }

        [TestMethod]
        public void B_Create_FilterColumns__Operators_Entity()
        {
            string response = null;
            response = DAL_SQL_Server.Factory.CREATE_ENTITY.CreationFactory.Create(ProjectPath, NameSpace, EntityName, listMyColumns);
            Console.WriteLine(response);
        }

        [TestMethod]
        public void CC_Create_FilterColumns__Operators_FileTypeFactory()
        {
            string response = null;
            response = DAL_SQL_Server.Factory.CREATE_ENTITY.CreationFactory.CreateFileTypeFactory(ProjectPath, NameSpace, EntityName, FactoryName, TableName, listMyColumns);
            Console.WriteLine(response);
        }

        [TestMethod]
        public void EE_Create_FilterColumns__Operators_FileFactory()
        {
            string response = null;
            response = DAL_SQL_Server.Factory.CREATE_ENTITY.CreationFactory.CreateFileFactory(ProjectPath, NameSpace, EntityName, FactoryName);
            Console.WriteLine(response);
        }

        #endregion
    }
}
