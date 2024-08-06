using System.ComponentModel;
using System.Reflection;
using System.Xml;
using System.Xml.Serialization;

namespace RealEstateTransactionsToCSV
{
    /// <summary>
    /// The program converts XML files (real_estate_transactions) to CSV files
    /// </summary>
    internal class Program
    {
        /// <summary>
        /// Start program
        /// </summary>
        /// <param name="args"></param>
        static void Main()
        {
            string curDir = Directory.GetCurrentDirectory();
            string xmlDir = Path.Combine(curDir, "XML");
            string outDir = Path.Combine(curDir, "OUT");

            if (!Directory.Exists(xmlDir)) Directory.CreateDirectory(xmlDir);
            if (!Directory.Exists(outDir)) Directory.CreateDirectory(outDir);

            Console.WriteLine($"1. Положите XML-файлы в папку {xmlDir}.");
            Console.WriteLine($"2. Нажмите ENTER для продолжения.");
            Console.ReadKey();

            var files = Directory.GetFiles(xmlDir, "*.xml");
            if (files.Length == 0)
            {
                Console.WriteLine($"\nОшибка! XML-файлов не обнаружено");
                Console.ReadKey();
                return;
            }

            Console.WriteLine($"\nКонвертируем файлы:");
            if (!XMLConverter(files, outDir)) return;


            Console.WriteLine("Ready! Press key Enter ...");
            Console.ReadKey();
        }

        /// <summary>
        /// XMLConverter converts XML files to CSV files
        /// </summary>
        private static bool XMLConverter(string[] files, string outDir)
        {
            // Время создания файлов
            string dateFile = DateTime.Now.ToString("yyyy-MM-dd HH-mm-ss");
            // Сделки на основании договоров купли-продажи
            string csvSaleContracts = Path.Combine(outDir, $"Сделки на основании договоров купли-продажи {dateFile}.csv");
            // Сделки на основании договоров аренды
            string csvLeaseContracts = Path.Combine(outDir, $"Сделки на основании договоров аренды {dateFile}.csv");
            // Сделки на основании договоров ипотеки
            string csvMortgageContracts = Path.Combine(outDir, $"Сделки на основании договоров ипотеки {dateFile}.csv");
            // Сделки на основании договора участия в долевом строительстве
            string csvSharedConstructionContracts = Path.Combine(outDir, $"Сделки на основании договора участия в долевом строительстве {dateFile}.csv");

            try
            {
                XmlSerializer xmlSerializer = new XmlSerializer(typeof(RealEstateTransactionsV01));
                foreach (string file in files)
                {
                    Console.WriteLine($"  {file} ...");
                    var resultConvert = (RealEstateTransactionsV01?)xmlSerializer.Deserialize(new XmlTextReader(file));

                    SaveSaleContracts(ref resultConvert, csvSaleContracts);
                    SaveLeaseContracts(ref resultConvert, csvLeaseContracts);
                    SaveMortgageContracts(ref resultConvert, csvMortgageContracts);
                    SaveSharedConstructionContracts(ref resultConvert, csvSharedConstructionContracts);

                    Console.WriteLine($"  ... ok");
                }
            }
            catch(Exception ex)
            {
                Console.WriteLine($"  ... Ошибка! {ex.Message}");
                Console.ReadKey();
                return false;
            }

            return true;
        }

        /// <summary>
        /// SaveSharedConstructionContracts saves the SharedConstructionContract result of convert into CSV file
        /// </summary>
        private static void SaveSharedConstructionContracts(ref RealEstateTransactionsV01? result, string csvFile)
        {
            if (result?.Transactions?.SharedConstructionContracts == null || result.Transactions.SharedConstructionContracts.Count == 0) return;

            bool addHeader = !File.Exists(csvFile);
            using StreamWriter stream = new StreamWriter(csvFile, true, System.Text.Encoding.UTF8);

            if (addHeader)
                stream.WriteLine(
                    "Вид сделки;" +
                    "Дата государственной регистрации договора;" +
                    "Дата выдачи (подписания) документа-основания;" +
                    "Предмет сделки;" +
                    "Цена, определенная договором;" +
                    "Кадастровый номер"
                );

            foreach (var item in result.Transactions.SharedConstructionContracts)
            {
                stream.WriteLine(
                    $"{item.RegistrationType.GetAttributeValue<DescriptionAttribute, string>(x => x.Description)};" +
                    $"{item.RegistrationDate};" +
                    $"{item.DocumentDate};" +
                    $"{item.Subject.Replace("\n", " ").Replace(";", ".")};" +
                    $"{item.ObjectsPrice};" +
                    $"{item.Objects.Object.CadastralNumber}"
                );
            }
        }

        /// <summary>
        /// SaveMortgageContracts saves the MortgageContracts result of convert into CSV file
        /// </summary>
        private static void SaveMortgageContracts(ref RealEstateTransactionsV01? result, string csvFile)
        {
            if (result?.Transactions?.MortgageContracts == null || result.Transactions.MortgageContracts.Count == 0) return;

            bool addHeader = !File.Exists(csvFile);
            using StreamWriter stream = new StreamWriter(csvFile, true, System.Text.Encoding.UTF8);

            if (addHeader)
                stream.WriteLine(
                    "Вид зарегистрированного ограничения права;" +
                    "Дата государственной регистрации ипотеки;" +
                    "Дата выдачи (подписания) документа-основания;" +
                    "Дата возникновения ипотеки в соответствии с договором об ипотеке;" +
                    "Дата исполнения обязательства, обеспеченного залогом в соответствии с договором об ипотеке;" +
                    "Оценка предмета ипотеки;" +
                    "Кадастровый номер"
                );

            foreach (var item in result.Transactions.MortgageContracts)
            {
                foreach (var ob in item.Objects)
                {
                    stream.WriteLine(
                        $"{item.RegistrationType.GetAttributeValue<DescriptionAttribute, string>(x => x.Description)};" +
                        $"{item.RegistrationDate};" +
                        $"{item.DocumentDate};" +
                        $"{item.OccurenceDate};" +
                        $"{item.ExecutionObligationsDate};" +
                        $"{item.ObjectsPrice.Replace("\n", " ").Replace(";", ".")};" +
                        $"{ob.CadastralNumber}"
                    );
                }
            }
        }

        /// <summary>
        /// SaveLeaseContracts saves the LeaseContracts result of convert into CSV file
        /// </summary>
        private static void SaveLeaseContracts(ref RealEstateTransactionsV01? result, string csvFile)
        {
            if (result?.Transactions?.LeaseContracts == null || result.Transactions.LeaseContracts.Count == 0) return;

            bool addHeader = !File.Exists(csvFile);
            using StreamWriter stream = new StreamWriter(csvFile, true, System.Text.Encoding.UTF8);

            if (addHeader)
                stream.WriteLine(
                    "Вид зарегистрированного ограничения права;" +
                    "Дата государственной регистрации сделки;" +
                    "Дата выдачи (подписания) документа-основания;" +
                    "Арендная плата (Цена сделки);" +
                    "Дата начала аренды;" +
                    "Дата конца аренды;" +
                    "Продолжительность аренды;" +
                    "Кадастровый номер"
                );

            foreach (var item in result.Transactions.LeaseContracts)
            {
                foreach (var ob in item.Objects)
                {
                    stream.WriteLine(
                        $"{item.RegistrationType.GetAttributeValue<DescriptionAttribute, string>(x => x.Description)};" +
                        $"{item.RegistrationDate};" +
                        $"{item.DocumentDate};" +
                        $"{item.Price};" +
                        $"{item.StartDate};" +
                        $"{item.EndDate};" +
                        $"{item.Duration};" +
                        $"{ob.CadastralNumber}"
                    );
                }
            }
        }

        /// <summary>
        /// SaveSaleContracts saves the SaleContracts result of convert into CSV file
        /// </summary>
        private static void SaveSaleContracts(ref RealEstateTransactionsV01? result, string csvFile)
        {
            if (result?.Transactions?.SaleContracts == null || result.Transactions.SaleContracts.Count == 0) return;
            
            bool addHeader = !File.Exists(csvFile);
            using StreamWriter stream = new StreamWriter(csvFile, true, System.Text.Encoding.UTF8);

            if (addHeader) 
                stream.WriteLine(
                    "Вид зарегистрированного права;" +
                    "Дата государственной регистрации права;" +
                    "Дата выдачи (подписания) документа-основания;" +
                    "Цена сделки по договору;" +
                    "Кадастровый номер;" +
                    "Цена по договору;" +
                    "Доля. Числитель;" +
                    "Доля. Знаменатель;" +
                    "Цена доли в праве по договору;" +
                    "Размер приобретаемой(-ых) доли(-ей)"
                );

            foreach (var item in result.Transactions.SaleContracts)
            {
                foreach (var ob in item.Objects.Object)
                {
                    stream.WriteLine(
                        $"{item.RegistrationType.GetAttributeValue<DescriptionAttribute, string>(x => x.Description)};" +
                        $"{item.RegistrationDate};" +
                        $"{item.DocumentDate};" +
                        $"{item.Price};" +
                        $"{ob.CadastralNumber};" +
                        $"{ob.Price};" +
                        $"{ob.PartNumerator};" +
                        $"{ob.PartDenomenator};" +
                        $"{ob.PartRightPrice};" +
                        $"{ob.ShareDescription}"
                    );
                }
            }
        }
    }

    static class EnumExpected
    {
        public static Expected GetAttributeValue<T, Expected>(this Enum enumeration, Func<T, Expected> expression) where T : Attribute
        {
            T attribute =
                enumeration
                    .GetType()
                    .GetMember(enumeration.ToString())
                    .Where(member => member.MemberType == MemberTypes.Field)
                    .FirstOrDefault()
                    .GetCustomAttributes(typeof(T), false)
                    .Cast<T>()
                    .SingleOrDefault();

            if (attribute == null)
                return default(Expected);

            return expression(attribute);
        }
    }
}
