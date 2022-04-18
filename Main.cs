using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CreateRoomsAndTags
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {

            Document doc = commandData.Application.ActiveUIDocument.Document;

            try
            {
                //создание списка уровней в модели
                List<Level> levelList = new FilteredElementCollector(doc)
               .OfClass(typeof(Level))
               .OfType<Level>()
               .ToList();

                //поиск необходимого типа марки помещения в модели
                FamilySymbol roomTagType = FindRoomTagType(doc);

                if (roomTagType == null)
                {
                    TaskDialog.Show("Ошибка", "Не найдено семейство \"Уровень_номер помещения\", загрузите семейство из файла");

                    //загрузка отсутствующего семейства марки помещения в модель
                    LoadFamily(doc);

                    roomTagType = FindRoomTagType(doc);

                    if (roomTagType == null)
                    {
                        TaskDialog.Show("Ошибка", "Файл семейства не выбран или некорректен");
                        return Result.Cancelled;
                    }
                }

                //вактивация семейства марки помещения
                Transaction transaction0 = new Transaction(doc);
                transaction0.Start("Активация семейства марки помещений");
                if (!roomTagType.IsActive)
                    roomTagType.Activate();
                transaction0.Commit();

                //создание помещений
                CreateRooms(doc, levelList);
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Ошибка", ex.Message);
                return Result.Failed;               
            }            

            TaskDialog.Show("Сообщение", "Создание помещений успешно завершено");
            return Result.Succeeded;
        }

        private static void CreateRooms(Document doc, List<Level> levelList)    //метод создающий помещения по уровням в модели
        {
            Transaction transaction = new Transaction(doc);
            transaction.Start("Создание помещений");
            if (levelList.Count() > 0)
            {
                foreach (Level level in levelList)
                {
                    //получаем план топологии выбранного уровня
                    PlanTopology planTopology = doc.get_PlanTopology(level);
                    if (planTopology != null)
                    {
                        //перебор замкнутых областей плана
                        foreach (PlanCircuit planCircuit in planTopology.Circuits)
                        {
                            //если в замкнутой области нет созданного помещения, создаём
                            if (!planCircuit.IsRoomLocated)
                            {                                
                                Room room = doc.Create.NewRoom(null, planCircuit);                                
                            }
                        }
                    }
                }
            }
            transaction.Commit();
        }

        private static FamilySymbol FindRoomTagType(Document doc)   //метод для поиска в модели необходимого для плагина семейства
        {
            return new FilteredElementCollector(doc)
                          .OfCategory(BuiltInCategory.OST_RoomTags)
                          .OfType<FamilySymbol>()
                          .Where(x => x.FamilyName.Equals("Уровень_номер помещения"))
                          .FirstOrDefault();
        }

        private void LoadFamily(Document doc)   //загрузка в модель необходимого семейства
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            //определение каталога по умолчанию (рабочий стол)
            openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop); 
            //фильтр для файлов rfa
            openFileDialog.Filter = "RFA files(*.rfa)|*.rfa";   

            //определение пути загружаемому файлу
            string filePath = string.Empty; 
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                filePath = openFileDialog.FileName;
            }

            //прерывание операции если путь пустой
            if (string.IsNullOrEmpty(filePath)) 
                return;

            Transaction transaction = new Transaction(doc);
            transaction.Start("Загрузка семейства марки помещений");
            doc.LoadFamily(filePath);
            transaction.Commit();
        }
    }
}
