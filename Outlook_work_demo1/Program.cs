using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Reflection;
using Microsoft.Office.Interop.Outlook;

namespace Outlook_work_demo1
{
    class Program
    {
        class FileM 
        {
            public static void MoveReplaceFile(string sourceFileName, string destFileName) // метод одназначного переноса файла
            {
                
                    //first, delete target file if exists, as File.Move() does not support overwrite
                    if (File.Exists(destFileName))
                    {
                        File.Delete(destFileName);// удаляем из папки куда копируется файл 
                    }

                    File.Move(sourceFileName, destFileName); // переместили
                if (File.Exists(destFileName))// если файл переместился то проверяем адплился ли он из основной папки
                {
                    if (File.Exists(sourceFileName))
                    {
                        File.Delete(sourceFileName);
                    }
                }
            }
        }
        static void Main(string[] args)
        {
            //Outlook.Application application; //
            Outlook._NameSpace nameSpace;
            Outlook.MAPIFolder folderInbox;
            //////////////////////////////////
            //Microsoft.Office.Interop.Outlook._Folders oFolders;
            ///////////////1 модуль отправки файлов///////////////////
            string[] dirs;
            dirs = Directory.GetFiles("C://OUT//"); //дирректория - откуда берём файл
            String dirs1 = Directory.GetFiles("C://OUT//").Length.ToString(); //если в папке out что то есть то отправляем письмо

            for (int i = 0; i <= dirs.Length + 1; i++)//для отправки 1 письмо + 1 файл
            {

                if (Convert.ToInt32(dirs1) > 0)
                {

                    //string dirArch = "C://OUT//ARCH//";
                    try
                    {
                        //тело отправляемого сообщения


                        Outlook._Application _app = new Outlook.Application();
                        Outlook.MailItem mail = (Outlook.MailItem)_app.CreateItem(Outlook.OlItemType.olMailItem);
                        mail.To = "maxim090491@outlook.com"; //"maximusvictor@mail.ru";
                        mail.Subject = "test1567";
                        mail.Body = "This is test message1156711";
                        mail.Importance = Outlook.OlImportance.olImportanceNormal;


                        foreach (string s in dirs)
                        {


                            try
                            {

                                Console.WriteLine(s);
                                mail.Attachments.Add(s); //добавляем вложения в письмо - возможно добавить сразу все вложения
                                File.Move(s, s.Replace("OUT", "ARCH")); //После успешного прикрепления файлой, переносим их в папку C:\ARCH
                                ((Outlook._MailItem)mail).Send();//отправляем письмо
                                Console.WriteLine("Файл отправлен в ТК");// если все хорошо, то выдаём в консоль сообщение
                                break;// выхожу из цикла, что бы реализовать фичу 1 фал = 1 письмо, и так сойдёт :) рефакторинг потом
                            }
                            catch (System.Exception ex)
                            {

                                Console.WriteLine(ex.Message);
                                Console.WriteLine("С таким именем файла уже отправлялось\n" + "Перемещаю файл в папку BadPack");
                                FileM.MoveReplaceFile(s, s.Replace("OUT", "BadPack"));
                                System.Threading.Thread.Sleep(1000);
                                Console.ReadKey();

                            }


                        }





                        _app = null; //убиваем ссылки на экземпляр класса
                        mail = null;

                        //Console.Read();

                    }
                    catch
                    {
                        Console.WriteLine("Что то не так блять"); // если какае то бага то пишем это сообщение
                        Console.WriteLine("Переходим к приёму файлов нажми любую клавишу");
                        System.Threading.Thread.Sleep(1000);
                        goto importStart; //если модуль отправки не отработал то программа переходит к модулю чтения сообщений

                    }
                    ////////////////////////////Чтение сообщений и копирование вложений///////////////////////////////////////////
                }
                else { Console.WriteLine("Файлов для отправки нет"); break; } //если нет файлов во вложении то выходим из цикла
                Console.ReadKey();
                dirs = Directory.GetFiles("C://OUT//"); //присваивание переменной с каждой итерацией цикла для обновления данных в массиве
                System.Threading.Thread.Sleep(1000);
            }
            importStart:
                Outlook.Application oApp = new Outlook.Application();// создали новый экземпляр
                Outlook.NameSpace oNS = oApp.GetNamespace("MAPI");

                //Это кусорк для чтения входящих сообщений для теста и понимания как это работает(раюотаем из папки входящие)
                oNS.Logon(Missing.Value, Missing.Value, false, true);
                Outlook.MAPIFolder oInbox = oNS.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);

                for (int x = 1; x <= oInbox.Items.Count; x++)
                {
                    if (oInbox.Items[x] is MailItem)
                    {
                        //Выводим Имя отправителя 
                        Console.WriteLine(oInbox.Items[x].SenderName + "\n" + "--------------------------------------------" + "\n");

                    }
                }
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


                Outlook.Items unreadItems = processedFolder.Items.Restrict("[Unread]=true");// работаем с данной папкой и ищем не прочитанные файлы
                Console.WriteLine("Непрочитанных писем = " + unreadItems.Count.ToString()); // выдать в консоль колличество непрочитанных сообщений

                //////////здесь вываодим в консоль все данные письма можно заккоментить/////////////////////////
                StringBuilder str = new StringBuilder(); //читаем текст и данные письма выыодим в консоль не обязательно, но можно оставить для логирования
                try
                {
                DateTime sysdate_ = DateTime.Now;
                foreach (Outlook.MailItem mailItem in unreadItems)
                    {
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

                    }
                }
                catch (System.Exception ex) // если письмо бракованное то пробуем ещё раз
                {
                    //System.Threading.Thread.Sleep(300); //пауза

                    //goto repeatReadMail;

                    string errorInfo = (string)ex.Message.Substring(0, 11);
                    Console.WriteLine(errorInfo);// вывести любую ошибку
                    if (errorInfo == "Cannot save")
                    {
                        Console.WriteLine(@"Create Folder C:\TestFileSave");
                    }
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
                Outlook.Items inBoxItems = processedFolder.Items.Restrict("[Unread]=true");

            Outlook.MailItem newEmail = null;
                try
                {
                    foreach (object collectionItem in inBoxItems)
                    {
                        newEmail = collectionItem as Outlook.MailItem;
                        DateTime sysdate_ = DateTime.Now; //
                    //mailItem.ReceivedTime.Date.ToShortDateString() + sysdate_.ToShortDateString();
                    if (newEmail != null && newEmail.ReceivedTime.Date.ToShortDateString() == sysdate_.ToShortDateString())
                        {
                        if (newEmail.Attachments.Count > 0)
                            {
                            for (int i = 1; i <= newEmail.Attachments.Count; i++)
                                {
                                string[] dirsIN = Directory.GetFiles("C://OUT//");//создадим перепенную с именами файлов в папке IN или IN_ARCH
                                //if ()
                                //    {
                                        newEmail.Attachments[i].SaveAsFile(@"C:\IN\" + newEmail.Attachments[i].FileName);
                                        Console.WriteLine(@"Файл сохранён с названием: " + newEmail.Attachments[i].FileName + " По пути:" + @"C:\IN\"); // выводим инфу об успешно копировании файла

                                        if (newEmail.UnRead) // пометить письмо как прочитанное
                                        {
                                            newEmail.UnRead = false;
                                            newEmail.Save();
                                        }
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
                oInbox = null;
                oNS = null;
                oApp = null;
                //Console.ReadKey(); //Удалить все чтения ввода с консоли, для автоматической работы
            System.Threading.Thread.Sleep(1000);

        }
        }
    }

