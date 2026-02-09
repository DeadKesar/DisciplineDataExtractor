using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using System;
using DisciplineDataExtractor.Word;

namespace DisciplineDataExtractor.Word.Helpers
{
    public static class Ooxml
    {
        public static IEnumerable<T> FindElementsByBookmark<T>(BookmarkStart bookmarkStart, uint outerLevels, WordprocessingDocument doc) where T : OpenXmlElement
        {
            var elements = new List<T>();
            var elem = bookmarkStart.NextSibling();

            if (elem == null)
            {
                var temp = FindElementsByBookmark2<T>(doc, bookmarkStart);
                return temp;
            }


            while (elem != null)
            {
                //Проверка самого элемента
                switch (elem)
                {
                    case BookmarkEnd el when el.Id == bookmarkStart.Id:
                        return elements;
                    case T element:
                        elements.Add(element);
                        break;
                }
                //Проверка всех элементов "под" самим элементом
                foreach (var node in elem.Descendants())
                {
                    switch (node)
                    {
                        case BookmarkEnd end when end.Id == bookmarkStart.Id:
                            return elements;
                        case T n:
                            elements.Add(n);
                            continue;
                    }
                }

                var next = elem.NextSibling();

                if (!(next is null))
                    elem = next;
                else if (outerLevels > 0)
                {
                    elem = elem.Parent;
                    outerLevels--;
                }
                else return elements;
            }

            return elements;
        }

        public static IDictionary<string, BookmarkStart> GetBookmarks(WordprocessingDocument doc, string bookmarkStartName)
        {
            var bookmarkMap = new Dictionary<string, BookmarkStart>();

            foreach (var bookmarkStart in doc.MainDocumentPart.Document.Body.Descendants<BookmarkStart>())
                if (bookmarkStart.Name.ToString().StartsWith(bookmarkStartName)) //Если сделать без идентификатора, то будут коллизии
                    bookmarkMap[bookmarkStart.Name.ToString().Substring(8)] = bookmarkStart; //Но Autofill отсекается

            return bookmarkMap;
        }

        public static void SaveDoc(OpenXmlPackage doc, string dir, string name)
        {
            //var regexSearch = new string(Path.GetInvalidFileNameChars()) + new string(Path.GetInvalidPathChars());
            name = RegexPatterns.InvalidChars.Replace(name, "");

            if (!Directory.Exists(dir))
                Directory.CreateDirectory(dir);

            string path = Path.Combine($"{dir}/{name}.docx");
            using (var stream = new FileStream(path, FileMode.Create, FileAccess.ReadWrite))
            {
                // Метод Create возвращает новый пакет, а затем данные копируются
                var newDoc = (doc as WordprocessingDocument).Clone(stream);
                newDoc.Save(); // Сохраняем документ
            }
        }

        public static IEnumerable<T> FindElementsByBookmark2<T>(WordprocessingDocument doc, BookmarkStart bookmarkStart) where T : OpenXmlElement
        {
            var body = doc.MainDocumentPart.Document.Body;
            var elements = new List<T>();

            // Ищем соответствующий BookmarkEnd по ID
            var bookmarkStartTemp = body.Descendants<BookmarkStart>()
                                  .FirstOrDefault(bs => bs.Id == bookmarkStart.Id);

            var bookmarkEnd = body.Descendants<BookmarkEnd>()
                                  .FirstOrDefault(be => be.Id == bookmarkStart.Id);
            if (bookmarkStartTemp == null)
                throw new Exception($"Не найдено начало закладки для ID {bookmarkStart.Id}");
            if (bookmarkEnd == null)
                throw new Exception($"Не найден конец закладки для ID {bookmarkStart.Id}");

            // Добавляем все элементы между BookmarkStart и BookmarkEnd
            bool isInsideBookmark = false;

            foreach (var element in body.Descendants())
            {
                // Начинаем сбор, как только найдём BookmarkStart
                if (element == bookmarkStartTemp)
                {
                    isInsideBookmark = true;
                    continue;
                }

                // Прекращаем сбор при нахождении BookmarkEnd
                if (element == bookmarkEnd)
                    break;

                if (isInsideBookmark)
                    if (element is T typedElement)
                    {
                        elements.Add(typedElement);
                    }
            }

            return elements;
        }



        /*public static IEnumerable<OpenXmlElement> FindElementsByBookmark2(BookmarkStart bookmarkStart)
        {
            // Получаем родительский контейнер
            var parent = bookmarkStart.Parent;
            var elements = new List<OpenXmlElement>();

            // Ищем все элементы после BookmarkStart
            var found = false;

            foreach (var child in parent.ChildElements)
            {
                if (child == bookmarkStart)
                {
                    found = true;
                    continue;
                }

                if (found)
                {
                    // Если встретили BookmarkEnd — завершаем поиск
                    if (child is BookmarkEnd bookmarkEnd && bookmarkEnd.Id == bookmarkStart.Id)
                        break;

                    elements.Add(child);
                }
            }

            return elements;
        }*/

    }
}