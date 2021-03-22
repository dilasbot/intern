using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;       //microsoft Excel 14 object in references-> COM tab
using System.Data;

namespace system
{
    public class Program
    {
        static void Main()
        {
            
            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"C:\Users\Hakan\Desktop\DÝLASUDE\StudentSync\di_db.xlsx");
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;

            List<Student> fromOtomasyon = new List<Student>();

            for (int i = 2; i <= rowCount; i++)
            {
                Student std = new Student
                {
                    EmailAddress = " "
                };
                for (int j = 1; j < 31; j++)
                {
                    switch (j)
                    {
                        case 1:
                            std.Id = xlRange.Cells[i, j].Value2.ToString();
                            break;
                        case 2:
                            std.Name = xlRange.Cells[i, j].Value2.ToString();
                            break;
                        case 3:
                            std.Surname = xlRange.Cells[i, j].Value2.ToString();
                            break;
                        case 4:
                            if (xlRange.Cells[i, j].Value2.ToString() == "M")
                            {
                                std.Sex = "Erkek";
                                break;
                            }
                            else
                            {
                                std.Sex = "Kadýn";
                                break;
                            }
                        case 5:
                            std.AcceptanceType = xlRange.Cells[i, j].Value2.ToString();
                            break;
                        case 7:
                            std.InstituteText = xlRange.Cells[i, j].Value2.ToString();
                            break;
                        case 8:
                            std.DepartmentText = xlRange.Cells[i, j].Value2.ToString();
                            break;
                        case 9:
                            std.ProgramText = xlRange.Cells[i, j].Value2.ToString();
                            break;
                        case 10:
                            if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                            {
                                std.NationalityText = xlRange.Cells[i, j].Value2.ToString();
                                break;
                            }
                            else
                                std.NationalityText = "Türkiye";
                            break;
                        case 11:
                            if (xlRange.Cells[i, j].Value2.ToString() == "Aktif")
                            {
                                std.Status = true;
                                break;
                            }
                            else
                            {
                                std.Status = false;
                                break;
                            }
                        case 13:
                            if (xlRange.Cells[i, j].Value2.ToString() == "YÜKSEK LÝSANS")
                            {
                                std.CycleID = 3;
                                break;
                            }
                            else if (xlRange.Cells[i, j].Value2.ToString() == "LÝSANS")
                            {
                                std.CycleID = 1;
                                break;
                            }
                            if (xlRange.Cells[i, j].Value2.ToString() == "DOKTORA")
                            {
                                std.CycleID = 4;
                                break;
                            }
                            else if (xlRange.Cells[i, j].Value2.ToString() == "SY")
                            {
                                std.CycleID = 5;
                                break;
                            }
                            else
                            {
                                std.CycleID = 2;
                                break;
                            }

                        case 14:
                            if (xlRange.Cells[i, j].Value2.ToString() == "HAZIRLIK")
                            {
                                std.LevelID = 1;
                                break;
                            }
                            else
                                break;
                        case 15:
                            string day = (xlRange.Cells[i, j].Value2.ToString());
                            double date = double.Parse(day);
                            std.RegistrationDate = DateTime.FromOADate(date);
                            break;
                        case 17:
                            if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                            {
                                std.PassedPreparationClass = true;
                                break;
                            }
                            else
                                std.PassedPreparationClass = false;
                            break;
                        case 20:
                            if (xlRange.Cells[i, j].Value2.ToString() == "-")
                            {
                                std.RegisteredTermCount = "0";
                                break;
                            }
                            else
                                std.RegisteredTermCount = xlRange.Cells[i, j].Value2.ToString();
                            break;
                        case 25:
                            if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                            {
                                std.CoAdvisor = xlRange.Cells[i, j].Value2.ToString();
                                break;
                            }
                            else
                                break;
                        case 26:
                            if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                            {
                                std.AdvisorText = xlRange.Cells[i, j].Value2.ToString();
                                break;
                            }
                            else
                                break;
                        case 27:
                            if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                            {
                                std.PhoneNumber = xlRange.Cells[i, j].Value2.ToString();
                                break;
                            }
                            else
                                break;
                        case 28:
                            if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                            {
                                std.EmailAddress = xlRange.Cells[i, j].Value2.ToString();
                                break;
                            }
                            else
                                break;

                        case 6:
                        case 12:
                        case 16:
                        case 18:
                        case 19:
                            break;
                        case 21:
                            if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                            {
                                std.LevelID = 5;
                                break;
                            }
                            else
                                break;
                        case 22:
                            if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                            {
                                std.LevelID = 3;
                                break;
                            }
                            else
                                break;
                        case 23:
                            if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                            {
                                std.LevelID = 4;
                                break;
                            }
                            else
                                break;
                        case 24:
                            if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                            {
                                std.LevelID = 2;
                                break;
                            }
                            else
                                break;
                        case 30:
                            if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                            {
                                std.ComprehensiveRegistrationCount = xlRange.Cells[i, j].Value2.ToString();
                                break;
                            }
                            else
                            {
                                std.ComprehensiveRegistrationCount = "0";
                                break;
                            }
                    }

                }
                if (xlRange.Cells[i, 29] != null && xlRange.Cells[i, 29].Value2 != null)
                {
                    if (std.EmailAddress == " ")
                    {
                        std.EmailAddress = xlRange.Cells[i, 29].Value2.ToString();
                    }
                }
                if (std.AdvisorText != null)
                {
                    std.CoAdvisor = std.CoAdvisor + ";" + std.AdvisorText;
                }

                FindProgDepartID(std);
                fromOtomasyon.Add(std);
            }
            GC.Collect();
            GC.WaitForPendingFinalizers();

            string connetionString, sql;
            SqlConnection cnn;
            SqlCommand command;
            SqlDataReader dataReader;
            List<Student> myDB = new List<Student>();

            connetionString = @"Data Source=.\SQLEXPRESS;Initial Catalog=BuroDB;User ID=id;Password=pass";
            sql = "SELECT StudentNumber, Name, Surname,Sex,AcceptanceType,ProgramText,DepartmentText,InstituteText,NationalityID,Status,CycleID,LevelID,RegistrationDate,StartDate,ForeignLanguageLastDate,GraduationDeadline,PassedPreparationClass,RegisteredTermCount,ComprehensiveRegistrationCount,AdvisorID,CoAdvisor,MilitaryServiceStatus,Homeaddress,PhoneNumber,EmailAddress FROM Students";

            cnn = new SqlConnection(connetionString);
            cnn.Open();
            command = new SqlCommand(sql, cnn);
            
            dataReader = command.ExecuteReader();

            while (dataReader.Read())
            {
                if (Convert.ToString(dataReader.GetValue(5)) != null)
                {
                    try
                    {
                        Student std2 = new Student
                        {
                            Id = Convert.ToString(dataReader.GetValue(0)),
                            Name = Convert.ToString(dataReader.GetValue(1)),
                            Surname = Convert.ToString(dataReader.GetValue(2)),
                            Sex = Convert.ToString(dataReader.GetValue(3)),
                            AcceptanceType = Convert.ToString(dataReader.GetValue(4)),
                            ProgramText = Convert.ToString(dataReader.GetValue(5)),
                            DepartmentText = Convert.ToString(dataReader.GetValue(6)),
                            InstituteText = Convert.ToString(dataReader.GetValue(7)),
                            NationalityText = Convert.ToString(dataReader.GetValue(8)),
                            Status = (dataReader.GetBoolean(9)),
                            CycleID = Convert.ToInt32(dataReader.GetValue(10)),
                            LevelID = Convert.ToInt32(dataReader.GetValue(11)),
                            RegistrationDate = Convert.ToDateTime(dataReader.GetValue(12)),
                            StartDate = Convert.ToDateTime(dataReader.GetValue(13)),
                            ForeignLanguageLastDate = Convert.ToString(dataReader.GetValue(14)),
                            GraduationDeadline = Convert.ToString(dataReader.GetValue(15)),
                            PassedPreparationClass = Convert.ToBoolean(dataReader.GetValue(16)),
                            RegisteredTermCount = Convert.ToString(dataReader.GetValue(17)),
                            ComprehensiveRegistrationCount = Convert.ToString(dataReader.GetValue(18)),
                            AdvisorText = Convert.ToString(dataReader.GetValue(19)),
                            CoAdvisor = Convert.ToString(dataReader.GetValue(20)),
                            MilitaryServiceStatus = Convert.ToString(dataReader.GetValue(21)),
                            Homeaddress = Convert.ToString(dataReader.GetValue(22)),
                            PhoneNumber = Convert.ToString(dataReader.GetValue(23)),
                            EmailAddress = Convert.ToString(dataReader.GetValue(24))
                        };
                        myDB.Add(std2);
                    }
                    catch { }
                }
            }
            List<Student> newStudents = new List<Student>(fromOtomasyon.Where(f => !myDB.Any(b => b.Id == f.Id)));
            List<Student> commonStudents = new List<Student>(from list1Item in fromOtomasyon
                                                             join list2Item in myDB on list1Item.Id equals list2Item.Id
                                                             where (list2Item != null)
                                                             select list1Item);
            List<Student> graduatedStudents = new List<Student>(myDB.Where(f => !fromOtomasyon.Any(b => b.Id == f.Id)));


            /*foreach (Student stu in graduatedStudents)
            {
                foreach (Student std in myDB)
                {
                    if (stu.Id == std.Id)
                    {
                        std.Type = "mezun";
                    }
                }
            }
            foreach (Student stu in newStudents)
            {
                foreach (Student std in fromOtomasyon)
                {
                    if (stu.Id == std.Id)
                    {
                        std.Type = "yeni";
                    }
                }
            }
            foreach (Student stu in commonStudents)
            {
                foreach (Student std in fromOtomasyon)
                {
                    if (stu.Id == std.Id)
                    {
                        std.Type = "halihazýrda";
                    }
                }
            }*/

            ChangeCommons(commonStudents);
            ChangeSurnames(myDB, fromOtomasyon);
            ChangeStatus(graduatedStudents);
            AddNewStudents(newStudents);

            dataReader.Close();
            cnn.Close();                     
            Console.WriteLine("bitti");
            Console.ReadLine();
            command.Dispose();           
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
        }
        public static void ChangeCommons(List<Student> commonStudents)
        {
            string updatesql;
            string connetionString = @"Data Source=.\SQLEXPRESS;Initial Catalog=DB;User ID=id;Password=pass";
            foreach (Student std in commonStudents)
            {
                updatesql = "UPDATE Students SET RegisteredTermCount = " + std.RegisteredTermCount + ", LevelID = " + std.LevelID + ", EmailAddress = '" + std.EmailAddress + "', CoAdvisor = '" + std.CoAdvisor + "' WHERE StudentNumber = '" + std.Id+"'";
                using (SqlConnection connection = new SqlConnection(connetionString))
                {
                    connection.Open();
                    using (SqlCommand commandd = new SqlCommand(updatesql, connection))
                    {
                        commandd.ExecuteNonQuery();
                        Console.WriteLine (std.Id + " " + std.Name + " " + std.Surname + " güncellendi.");
                        connection.Close();
                    }
                }
            }

        }
        public static void ChangeSurnames (List<Student> myDB, List<Student> fromOtomasyon)
        {
            foreach(Student std in myDB)
            {
                foreach(Student stdn in fromOtomasyon)
                {
                    if(std.Id==stdn.Id && std.Surname != stdn.Surname)
                    {
                        string updatesql;
                        string connetionString = @"Data Source=.\SQLEXPRESS;Initial Catalog=DB;User ID=id;Password=pass";
                        updatesql = "UPDATE Students SET Surname = " + stdn.Surname + " WHERE StudentNumber = '" + std.Id + "'";
                        using (SqlConnection connection = new SqlConnection(connetionString))
                        {
                            connection.Open();
                            using (SqlCommand commandd = new SqlCommand(updatesql, connection))
                            {
                                commandd.ExecuteNonQuery();
                                
                                connection.Close();
                            }
                        }
                        Console.WriteLine(std.Id + " " + std.Name + " " + stdn.Surname + " soyadý deðiþti.");

                    }
                }
            }
        }
        public static void FindProgDepartID(Student std)
        {
            switch(std.CycleID)
            {
                case 2:
                    if(std.ProgramText == "Yöneticiler için Ýþletme (MBA)")
                    {
                        std.ProgramID = 12;
                        std.DepartmentID = 7;
                        break;
                    }
                    else if (std.ProgramText == "Müzik (II.Öðretim Tezsiz)")
                    {
                        std.ProgramID = 16;
                        std.DepartmentID = 9;
                        break;
                    }
                    else if (std.ProgramText == "Giriþimcilik&Yenilik Yönetimi")
                    {
                        std.ProgramID = 6;
                        std.DepartmentID = 4;
                        break;
                    }
                    else if (std.ProgramText == "Ýþletme ve Teknoloji Yönetimi")
                    {
                        std.ProgramID = 14;
                        std.DepartmentID = 8;
                        break;
                    }
                    else
                    {
                        std.ProgramID = 0;
                        std.DepartmentID = 0;
                        break;
                    }
                case 3:
                    if (std.ProgramText == "Ýþletme (Tezsiz)")
                    {
                        std.ProgramID = 10;
                        std.DepartmentID = 7;
                        break;
                    }
                    else if (std.ProgramText == "Ýþletme")
                    {
                        std.ProgramID = 11;
                        std.DepartmentID = 7;
                        break;
                    }
                    else if (std.ProgramText == "Sanat Tarihi")
                    {
                        std.ProgramID = 23;
                        std.DepartmentID = 12;
                        break;
                    }
                    else if (std.ProgramText == "Müzikoloji")
                    {
                        std.ProgramID = 18;
                        std.DepartmentID = 10;
                        break;
                    }
                    else if (std.ProgramText == "Türk Müziði")
                    {
                        std.ProgramID = 0;
                        std.DepartmentID = 15;
                        break;
                    }
                    else if (std.ProgramText == "Ýç Mimari Tasarým Uluslararasý")
                    {
                        std.ProgramID = 7;
                        std.DepartmentID = 5;
                        break;
                    }
                    else if (std.ProgramText == "Siyaset Çalýþmalarý")
                    {
                        std.ProgramID = 25;
                        std.DepartmentID = 13;
                        break;
                    }
                    else if (std.ProgramText == "Çalgý-Ses Tezli")
                    {
                        std.ProgramID = 21;
                        std.DepartmentID = 11;
                        break;
                    }
                    else if (std.ProgramText == "Çalgý-Ses (Tezsiz)")
                    {
                        std.ProgramID = 21;
                        std.DepartmentID = 11;
                        break;
                    }
                    else if (std.ProgramText == "Geleneksel Danslar")
                    {
                        std.ProgramID = 22;
                        std.DepartmentID = 11;
                        break;
                    }
                    else if (std.ProgramText == "Bilim,Teknoloji ve Toplum")
                    {
                        std.ProgramID = 3;
                        std.DepartmentID = 2;
                        break;
                    }
                    else if (std.ProgramText == "Bilim ve Teknoloji Tarihi")
                    {
                        std.ProgramID = 1;
                        std.DepartmentID = 1;
                        break;
                    }
                    else
                    {
                        std.ProgramID = 0;
                        std.DepartmentID = 0;
                        break;
                    }
                case 4:
                    if (std.ProgramText == "Ýþletme")
                    {
                        std.ProgramID = 13;
                        std.DepartmentID = 7;
                        break;
                    }
                    else if (std.ProgramText == "Sanat Tarihi")
                    {
                        std.ProgramID = 24;
                        std.DepartmentID = 12;
                        break;
                    }
                    else if (std.ProgramText == "Müzik (MÝAM)")
                    {
                        std.ProgramID = 17;
                        std.DepartmentID = 9;
                        break;
                    }
                    else if (std.ProgramText == "Ýktisat")
                    {
                        std.ProgramID = 9;
                        std.DepartmentID = 6;
                        break;
                    }
                    else if (std.ProgramText == "Ýktisat (Ýngilizce)")
                    {
                        std.ProgramID = 9;
                        std.DepartmentID = 6;
                        break;
                    }
                    else if (std.ProgramText == "Müzikoloji ve Müzik Teorisi")
                    {
                        std.ProgramID = 20;
                        std.DepartmentID = 10;
                        break;
                    }
                    else if (std.ProgramText == "Siyasal&Toplumsal Düþünceler")
                    {
                        std.ProgramID = 26;
                        std.DepartmentID = 13;
                        break;
                    }
                    else
                    {
                        std.ProgramID = 0;
                        std.DepartmentID = 0;
                        break;
                    }
                case 5:
                    if (std.ProgramText == "Türk Sanat Müziði San.Yeterlik")
                    {
                        std.ProgramID = 0;
                        std.DepartmentID = 15;
                        break;
                    }
                    else if (std.ProgramText == "Türk Halk Müziði San.Yeterlik")
                    {
                        std.ProgramID = 0;
                        std.DepartmentID = 15;
                        break;
                    }
                    else
                    {
                        std.ProgramID = 0;
                        std.DepartmentID = 0;
                        break;
                    }
            }
                 
        }
        public static void ChangeStatus(List<Student> graduatedStudents)
        {
            string updatesql;
            string connetionString = @"Data Source=.\SQLEXPRESS;Initial Catalog=BuroDB;User ID=sa;Password=12345";
            foreach (Student std in graduatedStudents)
            {
                updatesql = "UPDATE Students SET Status = 0 WHERE StudentNumber = " + std.Id;               
                using (SqlConnection connection = new SqlConnection(connetionString))
                {
                    connection.Open();
                    using (SqlCommand commandd = new SqlCommand(updatesql, connection))
                    {
                        commandd.ExecuteNonQuery();
                        
                        connection.Close();
                    }
                }
                Console.WriteLine(std.Id + " " + std.Name + " " + std.Surname + " durumu pasif hale getirildi.");
            }
        }
        public static void AddNewStudents(List<Student> newStudents)
        {
            string insertsql;
           
            foreach (Student std in newStudents)
            {
                std.InstituteID = 1;
                std.NationalityID = 228;
                std.AdvisorID = 0;               
                std.ForeignLanguageLastDate = "2000-01-01 00:00:00.000";
                std.GraduationDeadline = "2000-01-01 00:00:00.000";
                insertsql = "INSERT INTO Students(StudentNumber, " +
                    "Name, " +
                    "Surname, " +
                    "Sex, " +
                    "AcceptanceType, " +
                    "ProgramID," +
                    "ProgramText," +
                    "DepartmentID," +
                    "DepartmentText," +
                    "InstituteID," +
                    "InstituteText," +
                    "NationalityID," +
                    "Status, " +
                    "CycleID, " +
                    "LevelID, " +
                    "RegistrationDate, " +
                    "StartDate, " +
                    "ForeignLanguageLastDate, " +
                    "GraduationDeadline, " +
                    "PassedPreparationClass, " +
                    "RegisteredTermCount, " +
                    "ComprehensiveRegistrationCount, " +
                    "AdvisorID," +
                    "CoAdvisor, " +
                    "PhoneNumber, " +
                    "EmailAddress) VALUES (" +
                    "@StudentNumber," +
                    "@Name," +
                    "@Surname," +
                    "@Sex," +
                    "@AcceptanceType," +
                    "@ProgramID," +
                    "@ProgramText," +
                    "@DepartmentID," +
                    "@DepartmentText," +
                    "@InstituteID," +
                    "@InstituteText," +
                    "@NationalityID," +
                    "@Status," +
                    "@CycleID," +
                    "@LevelID," +
                    "@RegistrationDate," +
                    "@StartDate," +
                    "@ForeignLanguageLastDate," +
                    "@GraduationDeadline," +
                    "@PassedPreparationClass," +
                    "@RegisteredTermCount," +
                    "@ComprehensiveRegistrationCount," +
                    "@AdvisorID," +
                    "@CoAdvisor," +
                    "@PhoneNumber," +
                    "@EmailAddress)";

                string connetionString = @"Data Source=.\SQLEXPRESS;Initial Catalog=DB;User ID=id;Password=pass";
                using (SqlConnection connection = new SqlConnection(connetionString))
                {
                    connection.Open();
                    using (SqlCommand commandd = new SqlCommand(insertsql, connection))
                    {
                        std.StartDate = std.RegistrationDate;

                        commandd.Parameters.AddWithValue("@StudentNumber", std.Id);
                        commandd.Parameters.AddWithValue("@Name", std.Name);
                        commandd.Parameters.AddWithValue("@Surname", std.Surname);
                        commandd.Parameters.AddWithValue("@Sex", std.Sex);
                        commandd.Parameters.AddWithValue("@AcceptanceType", std.AcceptanceType);
                        commandd.Parameters.AddWithValue("@Status", std.Status);
                        commandd.Parameters.AddWithValue("@CycleID", std.CycleID);
                        commandd.Parameters.AddWithValue("@LevelID", std.LevelID);
                        commandd.Parameters.AddWithValue("@ProgramID", std.ProgramID);
                        commandd.Parameters.AddWithValue("@RegistrationDate", std.RegistrationDate);
                        commandd.Parameters.AddWithValue("@StartDate", std.StartDate);
                        commandd.Parameters.AddWithValue("@DepartmentID", std.DepartmentID);
                        commandd.Parameters.AddWithValue("@InstituteID", std.InstituteID);
                        commandd.Parameters.AddWithValue("@InstituteText", std.InstituteText);
                        commandd.Parameters.AddWithValue("@NationalityID", std.NationalityID);
                        commandd.Parameters.AddWithValue("@AdvisorID", std.AdvisorID);
                        commandd.Parameters.AddWithValue("@ForeignLanguageLastDate", std.ForeignLanguageLastDate);
                        commandd.Parameters.AddWithValue("@GraduationDeadline", std.GraduationDeadline);
                        commandd.Parameters.AddWithValue("@PassedPreparationClass", std.PassedPreparationClass);
                        commandd.Parameters.AddWithValue("@RegisteredTermCount", std.RegisteredTermCount);
                        commandd.Parameters.AddWithValue("@ComprehensiveRegistrationCount", std.ComprehensiveRegistrationCount);
                        commandd.Parameters.AddWithValue("@CoAdvisor", std.CoAdvisor);
                        commandd.Parameters.AddWithValue("@PhoneNumber", std.PhoneNumber);
                        commandd.Parameters.AddWithValue("@EmailAddress", std.EmailAddress);
                        commandd.Parameters.AddWithValue("@ProgramText", std.ProgramText);
                        commandd.Parameters.AddWithValue("@DepartmentText", std.DepartmentText);
                        commandd.ExecuteNonQuery();                        
                        connection.Close();
                    }
                }
                Console.WriteLine(std.Id + " " + std.Name + " " + std.Surname + " sisteme eklendi.");
            }
        }
        public class Student
        {
            public string Id { get; set; }
            public string Name { get; set; }
            public string Surname { get; set; }
            public string Sex { get; set; }
            public string AcceptanceType { get; set; }
            public string ProgramText { get; set; }
            public string DepartmentText { get; set; }
            public string InstituteText { get; set; }
            public string NationalityText { get; set; }
            public bool Status { get; set; }
            public int CycleID { get; set; }
            public int LevelID { get; set; }
            public int ProgramID { get; set; }
            public int DepartmentID { get; set; }
            public int NationalityID { get; set; }
            public int InstituteID { get; set; }
            public int AdvisorID { get; set; }
            public DateTime RegistrationDate { get; set; }
            public DateTime StartDate { get; set; }
            public string ForeignLanguageLastDate { get; set; }
            public string GraduationDeadline { get; set; }
            public bool PassedPreparationClass { get; set; }
            public string RegisteredTermCount { get; set; }
            public string ComprehensiveRegistrationCount { get; set; }
            public string AdvisorText { get; set; }
            public string CoAdvisor { get; set; }
            public string MilitaryServiceStatus { get; set; }
            public string Homeaddress { get; set; }
            public string PhoneNumber { get; set; }
            public string EmailAddress { get; set; }
            public string Type { get; set; }
        }
    }
}

