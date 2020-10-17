using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Reflection;
using Microsoft.Office.Interop.Outlook;
using System.Runtime.Remoting.Messaging;
using System.Threading;

namespace Outlook_work_demo1
{
    class Program
    {
        
        static void Main(string[] args)
        {
        repiat:
            Console.Clear();
            Console.WriteLine("НАчинаю работу приложения");
            /************************переменные для логера**************************/
             const string Error_ = "Ошибка";
            const string Warning_ = "Предупреждение";
            const string Event_ = "Событие";
            /************************переменные для логера конец********************/
            //Logger.Run_();
            Logger.WriteLog(Event_, 0, "Старт");
            /**********Выполняем батники для отправки**********/
            /***********Переменные для батников*********/
            string batPuttyOut1 = @"C:\tkM\copyOUTputty.bat";
            string batPuttyOut2 = @"C:\tkM\copyOUTarch.bat";
            string batInPutty1 = @"C:\tkM\copyOUTarch.bat";
            string batInPutty2 = @"C:\tkM\copyOUTarch.bat";

            string pathBat = @"C:\tkM";

            string descrBatO1 = "fromPuttyToOut1";
            string descrBatO2 = "fromPuttyToOut2";
            string descrBatI3 = "fromPuttyToIN1";
            string descrBatI4 = "fromPuttyToIN2";


            //*************батники на отправку*********//
            FileM.ProcManage(batPuttyOut1, pathBat, descrBatO1);
            FileM.ProcManage(batPuttyOut2, pathBat, descrBatO2);

           
            //try 
            //{ 
            //System.Diagnostics.Process proc = new System.Diagnostics.Process();
            //proc.StartInfo.FileName = @"C:\tkM\copyOUTarch.bat";
            //proc.StartInfo.WorkingDirectory = @"C:\tkM";
            //proc.Start();

            //proc.WaitForExit();
            //    Console.WriteLine("Батники на отправку успешно отработали");
            //    Logger.WriteLog(Event_ , 0, "Батники на отправку успешно отработали");
            //    proc = null; //ufc
            //}
            //catch(System.Exception ex)
            //{

            //    Logger.WriteLog(Error_, 100 , "ошибка выолнения батников" +  Convert.ToString(ex.Message));
            //    Console.WriteLine("ошибка выолнения батников" + Convert.ToString(ex.Message));
            //}
            /********************/

            Logger.WriteLog(Event_, 0, "Выполняю модуль отправки писем"); Console.WriteLine("Выполняю модуль отправки писем");
            //Outlook.Application application; //
            Outlook._NameSpace nameSpace; 
            Outlook.MAPIFolder folderInbox;
            //////////////////////////////////
            //Microsoft.Office.Interop.Outlook._Folders oFolders;
           
            ///////////////1 модуль отправки файлов///////////////////
            string[] dirs;
            dirs = Directory.GetFiles(@"C:\tkM\OUT"); //дирректория - откуда берём файл
            String dirs1count = Directory.GetFiles(@"C:\tkM\OUT").Length.ToString(); //если в папке out что то есть то отправляем письмо
            
            for (int i = 0; i < dirs.Length; i++)//для отправки 1 письмо + 1 файл
            {
                Logger.WriteLog(Event_, 0, "Запускаю цикл обработки и отправки входящих файлов");

                if (Convert.ToInt32(dirs1count) > 0) 
                    
                {
                    Logger.WriteLog(Event_, 0, "Колличество файлов на отправку: " + dirs1count); Console.WriteLine("Выполняю модуль отправки писем");
                    //string dirArch = "C://OUT//ARCH//";
                    try
                    {
                        //тело отправляемого сообщения


                        Outlook._Application _app = new Outlook.Application();
                        Outlook.MailItem mail = (Outlook.MailItem)_app.CreateItem(Outlook.OlItemType.olMailItem);
                        mail.To = "maxim090491@outlook.com"; //"maximusvictor@mail.ru";
                        mail.Subject = "test1567"; // тема  
                        mail.Body = "This is test message1156711"; //текст письма
                        mail.Importance = Outlook.OlImportance.olImportanceNormal; //какае то нужная вещь


                        foreach (string s in dirs)
                        {


                            try
                            {

                                Console.WriteLine(s);
                                mail.Attachments.Add(s); //добавляем вложения в письмо - возможно добавить сразу все вложения
                                File.Move(s, s.Replace("OUT", "ARCH")); //После успешного прикрепления файла, переносим их в папку C:\ARCH
                                ((Outlook._MailItem)mail).Send();//отправляем письмо
                                Console.WriteLine("Файл" + s + "отправлен в ТК");// если все хорошо, то выдаём в консоль сообщение
                                Logger.WriteLog(Event_, 0, "Файл" + s + "отправлен в ТК");
                                break;// выхожу из цикла, что бы реализовать фичу 1 фал = 1 письмо, и так сойдёт :) рефакторинг потом
                                
                            }
                            catch (System.Exception ex)
                            {
                                Logger.WriteLog(Warning_, 201, "Системное исключение" + ex.Message);
                                Logger.WriteLog(Warning_, 200, "С таким именем файла уже отправлялось\n" + "Перемещаю файл" + s + "в папку Bad также прочти системное исключение" );
                                
                                Console.WriteLine("С таким именем файла уже отправлялось\n" + "Перемещаю файл в папку Bad");
                                FileM.MoveReplaceFile(s, s.Replace("OUT", "Bad"));//прописываю каталог C:\tkM\Bad так быстрее не люблю много переменных
                                System.Threading.Thread.Sleep(1000); //спим 1000 мс что б увидеть работу кода
                            }


                        }
                        Logger.WriteLog(Event_, 0, "успешно передан файл"  + dirs[i]);
                        _app = null; //убиваем ссылки на экземпляр класса
                        mail = null;

                        //Console.Read();

                    }
                    catch (System.Exception ex)
                    {

                        Logger.WriteLog(Error_, 111,  "Системная информация" + ex.Message);
                        Console.WriteLine("Что то пошло не так"); // если какае то бага то пишем это сообщение
                        Console.WriteLine("Не удалось отправить файлы, переходим к приёму файлов");
                        System.Threading.Thread.Sleep(1000);
                        Logger.WriteLog(Warning_, 300, "Переходим к модулю обработки входящих писем");
                        goto importStart; //если модуль отправки не отработал то программа переходит к модулю чтения сообщений

                    }
                    ////////////////////////////Чтение сообщений и копирование вложений///////////////////////////////////////////
                }
                else { Console.WriteLine("Файлов для отправки нет"); Logger.WriteLog(Event_, 0, "Файлов для отправки нет - Переходим к модулю обраюлтки входящих писем"); break; } //если нет файлов во вложении то выходим из цикла
                dirs1count = Directory.GetFiles(@"C:\tkM\OUT").Length.ToString();
                dirs = Directory.GetFiles(@"C:\tkM\OUT"); //переинициализируем переменную с каждым шагом цикла - так как файлов то меньше на 1
                System.Threading.Thread.Sleep(1000);
            }
            importStart:
            Logger.WriteLog(Event_, 0, "Обрабатываем входящие письма");

            Outlook.Application oApp = new Outlook.Application();// создали новый экземпляр
                Outlook.NameSpace oNS = oApp.GetNamespace("MAPI");

                //Это кусорк для чтения входящих сообщений для теста и понимания как это работает(раюотаем из папки входящие)
                //oNS.Logon(Missing.Value, Missing.Value, false, true);
                //Outlook.MAPIFolder oInbox = oNS.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);

                //for (int x = 1; x <= oInbox.Items.Count; x++)
                //{
                //    if (oInbox.Items[x] is MailItem)
                //    {
                //        //Выводим Имя отправителя 
                //        Console.WriteLine(oInbox.Items[x].SenderName + "\n" + "--------------------------------------------" + "\n");

                //    }
                //}
                nameSpace = oApp.GetNamespace("MAPI");
                object missingValue = System.Reflection.Missing.Value;
                folderInbox = nameSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox); //Outlook.OlDefaultFolders.olFolderInbox
                                                                                                  //количество не прочитанных писем в папке Входящие (Inbox)
                Outlook.MAPIFolder rootFolder = nameSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);//получаем доступ к папке входящие
                Outlook.MAPIFolder processedFolder = null; //processedFolder это экземпляр класса который выбирает нужную папку
                foreach (Outlook.MAPIFolder folder in rootFolder.Folders)// ищем циклом папку в OUTLOOK
                {
                    if (folder.Name == "IN") //берем из папки IN
                    {
                        processedFolder = folder; // выбираю папку с которой будем работать(в OUTLOOK нужно создать правило, что бы все вообщения от почты ТК переносились в папку IN, которую мы заранее создаём в OUTLOOK)
                        break;
                    }
                }


                Outlook.Items unreadItems = processedFolder.Items.Restrict("[Unread]=true");// работаем с данной папкой и ищем не прочитанные письма
                Console.WriteLine("Непрочитанных писем в папке in = " + unreadItems.Count.ToString()); // выдать в консоль колличество непрочитанных сообщений
            Logger.WriteLog(Event_, 0, "Непрочитанных писем в папке in = " + unreadItems.Count.ToString());
            //////////здесь вываодим в консоль все данные непрочитанные письма в папке in    можно заккоментить/////////////////////////
            StringBuilder str = new StringBuilder(); //читаем текст и данные письма выыодим в консоль не обязательно, но можно оставить для логирования
                try
                {
                DateTime sysdate_ = DateTime.Now;
                foreach (Outlook.MailItem mailItem in unreadItems)
                    {
                    Logger.WriteLog(Event_, 0, "Читаю письмо");
                    str.AppendLine("----------------------------------------------------");
                        str.AppendLine("SenderName: " + mailItem.SenderName);
                        str.AppendLine("To: " + mailItem.To);
                        str.AppendLine("CC: " + mailItem.CC);
                        str.AppendLine("Subject: " + mailItem.Subject);
                        str.AppendLine("Body: " + mailItem.Body);
                        str.AppendLine("CreationTime: " + mailItem.CreationTime);
                        str.AppendLine("ReceivedByName: " + mailItem.ReceivedByName);
                        str.AppendLine("ReceivedTime: " + mailItem.ReceivedTime);
                        str.AppendLine("UnRead: " + mailItem.UnRead);
                    str.AppendLine(mailItem.ReceivedTime.Date.ToShortDateString()  + sysdate_.ToShortDateString());
                    str.AppendLine("----------------------------------------------------");
                    Console.WriteLine(str.ToString());
                    Logger.WriteLog(Event_, 0, str.ToString());

                }
                }
                catch (System.Exception ex) // если письмо бракованное то пробуем ещё раз
                {
                //System.Threading.Thread.Sleep(300); //пауза

                //goto repeatReadMail;
                Logger.WriteLog(Error_, 102, "Ощибка сохранения письма или" + ex.Message);
                string errorInfo = (string)ex.Message.Substring(0, 11);
                    Console.WriteLine(errorInfo);// вывести любую ошибку
                    if (errorInfo == "Cannot save")
                    {
                        Console.WriteLine(@"Create Folder C:\TestFileSave");
                    }
                Logger.WriteLog(Event_, 0, "Перехожу к чтению вложений");
                goto repeatReadAndAddAttachmentIsMail; // перейти к чтению и вытаскиванию вложения в случае ошибки здесь
                }
            //foreach (Outlook.MailItem mail in unreadItems) //пометить как прочитанное реализация не удалять
            //{
            //    if (mail.UnRead)
            //    {
            //        mail.UnRead = false;
            //        mail.Save();
            //    }
            //}
            //////////////////Вытаскивание вложение и пометить письмо как прочитанное
            repeatReadAndAddAttachmentIsMail:
                Outlook.Items inBoxItems = processedFolder.Items.Restrict("[Unread]=true");//show unread message and inicialise varible

            Outlook.MailItem newEmail = null;
                try
                {
                    foreach (object collectionItem in inBoxItems)
                    {
                        newEmail = collectionItem as Outlook.MailItem;
                        DateTime sysdate_ = DateTime.Now; //SYSDATE
                    //mailItem.ReceivedTime.Date.ToShortDateString() + sysdate_.ToShortDateString();
                    if (newEmail != null /*&& newEmail.ReceivedTime.Date.ToShortDateString() == sysdate_.ToShortDateString()*/)//checj date of mail
                        {
                            if (newEmail.Attachments.Count > 0)
                            {
                            for (int i = 1; i <= newEmail.Attachments.Count; i++)
                                {
                                string fileName = newEmail.Attachments[i].FileName; Console.WriteLine(@"Имя файла:" + fileName);
                                string[] dirsaRCH = Directory.GetFiles(@"C:\tkM\ARCH\");//создадим перепенную с именами файлов в папке  ARCH что бы проверять файлы в архиве
                                if (FileM.ValueOf(dirsaRCH, fileName) != "noDuble")
                                {
                                    Console.WriteLine(@"Найдено совпаление  с файлом: " + fileName + "\n Список файлов в папке:");
                                    Logger.WriteLog(Warning_, 0, @"Найдено совпаление  с файлом: " + fileName + "не обратываем"); 

                                    //for (int i1 = 0; i1 < dirsaRCH.Length; i1++)
                                    //{
                                    //    Console.Write( dirsaRCH[i1] + " " + File.GetCreationTime(dirsaRCH[i1]) + " - не обратываем" + ";\n");
                                    //    Logger.WriteLog(Warning_, 205, dirsaRCH[i1] + " " + File.GetCreationTime(dirsaRCH[i1]) + " - не обратываем" + ";\n");
                                    //}

                                }
                                else 
                                {
                                    Console.WriteLine(@"нет совпадения файлов"); Logger.WriteLog(Event_, 0, @"Совпадений по имени не найдено");
                                    //Console.WriteLine(@"Имя файла:" + fileName);
                                    //for (int i1 = 0; i1 < dirsaRCH.Length; i1++)
                                    //{
                                    //    Console.Write("Имя файлов(а) в папке:" + dirsaRCH[i1]);
                                    //}
                                    newEmail.Attachments[i].SaveAsFile(@"C:\tkM\IN\" + newEmail.Attachments[i].FileName);
                                    Console.WriteLine(@"Файл сохранён с названием: " + newEmail.Attachments[i].FileName + " По пути:" + @"C:\tkM\IN\"); // выводим инфу об успешно копировании файла
                                    Logger.WriteLog(Event_, 0, @"Файл сохранён с названием: " + newEmail.Attachments[i].FileName + " По пути:" + @"C:\tkM\IN\");
                                }

                                //    {
                                Console.WriteLine(@"Обрабатываю письмо как прочитанное");
                                Logger.WriteLog(Event_, 0, @"Обрабатываю письмо как прочитанное");

                                System.Threading.Thread.Sleep(1000);
                                if (newEmail.UnRead) // пометить письмо как прочитанное
                                        {
                                            newEmail.UnRead = false;
                                            newEmail.Save();
                                        }

                                /********батники на IN************/
                                FileM.ProcManage(batInPutty1, pathBat, descrBatI3);
                                FileM.ProcManage(batInPutty2, pathBat, descrBatI4);
                                /********************/
                                string[] dirIN = Directory.GetFiles(@"C:\tkM\IN");
                                foreach(string d in dirIN)
                                {
                                    FileM.MoveReplaceFile(d, d.Replace("IN", "ARCH"));

                                }
                                    
                                Console.WriteLine(@"Завершаю работу");
                                
                                //}
                            }
                            }
                        
                        }
                    else { Console.WriteLine("Нет писем с вложением"); }
                    }
                }
                catch (System.Exception ex)
                {
                    string errorInfo = (string)ex.Message
                        .Substring(0, 11);
                    Console.WriteLine(errorInfo);
                    if (errorInfo == "Cannot save")
                    {
                        Console.WriteLine(@"Create Folder C:\IN");

                    }
                    
                }

                //////////выход из приложения
                oNS.Logoff();
                //oInbox = null;
                oNS = null;
                oApp = null;
            //Console.ReadKey(); //Удалить все чтения ввода с консоли, для автоматической работы
            Logger.Flush();
            Console.WriteLine($"Закончили обработку в {DateTime.Now} " +
                $"\nСледующий запуск приложения в {DateTime.Now.Add(TimeSpan.FromMinutes(5))} ");
            System.Threading.Thread.Sleep(15000);

           
            goto repiat;
            
        }
            class FileM
            {
            const string Error_ = "Ошибка";
            const string Warning_ = "Предупреждение";
            const string Event_ = "Событие";
            public static void MoveReplaceFile(string sourceFileName, string destFileName) // метод одназначного переноса файла
                {

                    //first, delete target file if exists, as File.Move() does not support overwrite
                    if (File.Exists(destFileName))
                    {
                        File.Delete(destFileName);// удаляем из папки куда копируется файл 
                    }

                    File.Move(sourceFileName, destFileName); // переместили
                    if (File.Exists(destFileName))// если файл переместился то проверяем удаплился ли он из основной папки
                    {
                        if (File.Exists(sourceFileName))
                        {
                            File.Delete(sourceFileName);
                        }
                    }
                }

             public static string ValueOf(string[] arrayS, string value ) //поиск названия в массиве
            {
                for(int i = 0; i < arrayS.Length; i++)
                {
                    if (Path.GetFileName(arrayS[i]) == value) //Класс и медод Path.GetFileName  вытаскивает из пути имя файла
                    {
                        return value;
                    }
                }
                return "noDuble";
            }
            public static string ProcManage(string procName, string pathDir, string descBat)
            {
                try
                {
                    Logger.WriteLog(Event_, 0, descBat);
                    System.Diagnostics.Process proc1 = new System.Diagnostics.Process();
                    proc1.StartInfo.FileName = procName;
                    proc1.StartInfo.WorkingDirectory = pathDir;
                    proc1.Start();

                    Logger.WriteLog(Warning_, 200, "Ожидаю выполнения:" + procName);
                    proc1.WaitForExit();

                    Console.WriteLine("Батник - " + descBat + " - успешно отработал");
                    Logger.WriteLog(Event_, 0, "Батник " + descBat + " - успешно отработал");
                    proc1 = null; //ufc
                }
                catch (System.Exception ex)
                {
                    Console.WriteLine("ошибка выполнения батников" + Convert.ToString(ex.Message));
                    Logger.WriteLog(Error_, 104, "ошибка выполнения батникa:" + procName + " " + Convert.ToString(ex.Message));
                    return "100";
                }
                return "0";
            }

          
            
            }
        }
    }

