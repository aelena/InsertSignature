using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using WORD = Microsoft.Office.Interop.Word;
using EXCEL = Microsoft.Office.Interop.Excel;
using System.Drawing;
using System.Globalization;
using System.Threading;

namespace Utils.Office
{
    public class OfficeSigner
    {

        /// <summary>
        /// Añade una firma escaneada en un documento 
        /// </summary>
        /// <param name="documentPath"></param>
        /// <param name="signatureFile"></param>
        public void InsertScannedSignature(string documentPath, string signatureFile, bool openDoc)
        {

            if (!(File.Exists(documentPath) && File.Exists(signatureFile)))
                throw new ArgumentException("Una de las rutas no es válida.");

            string _getDocExt = Path.GetExtension(documentPath).ToLowerInvariant().Substring(1);

            switch (Path.GetExtension(documentPath).ToLowerInvariant().Substring(1))
            {
                case "docx":
                case "doc":
                    this.SignWordDocument(documentPath, signatureFile, openDoc);
                    break;
                case "xlsx":
                case "xls":
                    this.SignExcelDocument(documentPath, signatureFile, openDoc);
                    break;
            }


        }   // cierre InsertScannedSignature

        private void SignWordDocument(string documentPath, string signatureFile, bool openWord)
        {
            WORD.ApplicationClass WordApp = new WORD.ApplicationClass();

            try
            {


                object _missing = System.Reflection.Missing.Value;
                object _docPath = documentPath;

                WORD.Document adoc = WordApp.Documents.Open(ref _docPath, ref _missing,
                    ref _missing, ref _missing, ref _missing,
                        ref _missing, ref _missing, ref _missing, ref _missing,
                            ref _missing, ref _missing, ref _missing, ref _missing,
                                ref _missing, ref _missing, ref _missing);

                object unit = 6;
                object move = 0;
                int _endKey = WordApp.Selection.EndKey(ref unit, ref move);
                // int _endOfDoc = WordApp.Selection.End;

                WordApp.Selection.InlineShapes.AddPicture(signatureFile, ref _missing, ref _missing, ref _missing);
                WordApp.Visible = openWord;

            }
            catch (Exception ex)
            {
                string _s = ex.ToString();
            }
            finally
            {
                WordApp = null;
            }

        }   // cierre SignWordDocument


        // ----------------------------------------------------------------------------------


        private void SignExcelDocument(string documentPath, string signatureFile, bool openDoc)
        {
            EXCEL.Application excelApp = new EXCEL.ApplicationClass();
            EXCEL.Workbook _workbook;
            CultureInfo _prevCulture = Thread.CurrentThread.CurrentCulture;

            // BUg de EXcel se soluciona con CUltureINfo VEr ENlace
            // http://support.microsoft.com/default.aspx?scid=kb;en-us;320369

            try
            {
                excelApp.UserControl = true;

                object _missing = System.Reflection.Missing.Value;
                object _docPath = documentPath;

                Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");

                _workbook = excelApp.Workbooks.Open(documentPath,
                     _missing, _missing, 5,
                         _missing, _missing, _missing, _missing,
                             _missing, _missing, _missing, _missing,
                                 _missing, _missing, _missing);

                EXCEL.Worksheet _sheet = (EXCEL.Worksheet)_workbook.Sheets[1];
                _sheet.Cells[15, 1] = DateTime.Now.TimeOfDay.ToString();

                EXCEL.Range _range = (EXCEL.Range)_sheet.Cells[17, 1];

                Image _img = Image.FromFile(signatureFile);
                // _range.set_Value(null, _img);
                _range.set_Item(17, 1, _img);
                _sheet.Paste(_range, signatureFile);
                _workbook.Save();

                excelApp.Visible = openDoc;

            }
            catch (Exception ex)
            {
                string _s = ex.ToString();
            }
            finally
            {
                excelApp.Quit();
                excelApp = null;
                _workbook = null;
                Thread.CurrentThread.CurrentCulture = _prevCulture;
            }

        }   // cierre SignExcelDocument


        // ----------------------------------------------------------------------------------


        public void InsertDigitalSignature(string documentPath, bool openDoc)
        {
            
            SPFile _spFile = properties.ListItem.File;  // obtener las propiedades a traves del wf
            // crear un nombre automatico para el fichero temporal
            string _fileTempName = System.Guid.NewGuid();
            //properties.ListItem[Strings.IDColumn].ToString() +
            //properties.ListItem[Strings.NameColumn].ToString();

            string _ext = Path.GetExtension(documentPath).ToLowerInvariant().Substring(1);

            if( _ext.Equals ("docx") || _ext.Equals ( "doc") )
            {
                // bajar el archivo
               byte[] _filebytes = _spFile.OpenBinary(SPOpenBinaryOptions.None);

                DirectoryInfo _tempDirectory = new DirectoryInfo(@"C:\Temp");
                if (!_tempDirectory.Exists)
                {
                    _tempDirectory.Create();
                }

                using (FileStream _tempFile = new FileStream(_tempDirectory + @"\" + _fileTempName,
                    FileMode.CreateNew))
                {
                    BinaryWriter bWriter = new BinaryWriter(_tempFile);
                    foreach (byte b in _filebytes)
                    {
                        bWriter.Write(b);
                    }
                    bWriter.Flush();
                    bWriter.Close();
                    _tempFile.Close();
                }

                bool _signatureAdded = AddDigitalSignature(_tempDirectory + @"\" + _fileTempName, properties);
                if (_signatureAdded)
                {
                    properties.ListItem.ParentList.ParentWeb.AllowUnsafeUpdates = true;
                    WriteFileAgain(properties, _fileTempName);
                    properties.ListItem.ParentList.ParentWeb.AllowUnsafeUpdates = false;

                }

            }   // cierre if

        }   // cierre InsertDigitalSignature


        // ----------------------------------------------------------------------------------


        /// <summary>
        /// Descarga un archivo residente en SharePoint a la ruta indicada.
        /// </summary>
        /// <param name="_spFile">Objecto SPFile</param>
        /// <param name="fullPath">Ruta y nombre completo con el que se va a guardar
        /// el archivo. De momento no contempla verificar rutas.</param>
        private bool DownloadFileFromMOSS(SPFile _spFile, string fullPath)
        {
            try
            {
                if (_spFile != null)
                {
                    // bajar el archivo
                    byte[] _filebytes = _spFile.OpenBinary(SPOpenBinaryOptions.None);
                    using (FileStream _tempFile = new FileStream(fullPath,
                        FileMode.CreateNew))
                    {
                        BinaryWriter bWriter = new BinaryWriter(_tempFile);
                        foreach (byte b in _filebytes)
                        {
                            bWriter.Write(b);
                        }
                        bWriter.Flush();
                        bWriter.Close();
                        _tempFile.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                // TODO: log error
                return false;
            }
            return true;


        }   // cierre DownloadFileFromMOSS


        // ----------------------------------------------------------------------------------


        private void UploadFile(string srcUrl, string _destSiteURL)
        {
            if (!File.Exists(srcUrl))
            {
                throw new ArgumentException(String.Format("{0} no existe", srcUrl), "srcUrl");
            }

            SPWeb site = new SPSite(_destSiteURL).OpenWeb();

            FileStream fStream = File.OpenRead(srcUrl);
            byte[] contents = new byte[fStream.Length];
            fStream.Read(contents, 0, (int)fStream.Length);
            fStream.Close();

            EnsureParentFolder(site, _destSiteURL);
            site.Files.Add(_destSiteURL, contents);

        }   // cierre UploadFile


        // ----------------------------------------------------------------------------------


        public string EnsureParentFolder(SPWeb parentSite, string _destSiteURL)
        {
            _destSiteURL = parentSite.GetFile(_destSiteURL).Url;

            int index = _destSiteURL.LastIndexOf("/");
            string parentFolderUrl = string.Empty;

            if (index > -1)
            {
                parentFolderUrl = _destSiteURL.Substring(0, index);

                SPFolder parentFolder = parentSite.GetFolder(parentFolderUrl);

                if (!parentFolder.Exists)
                {
                    SPFolder currentFolder = parentSite.RootFolder;

                    foreach (string folder in parentFolderUrl.Split('/'))
                    {
                        currentFolder = currentFolder.SubFolders.Add(folder);
                    }
                }
            }

            return parentFolderUrl;
        }


        // ----------------------------------------------------------------------------------


        private bool AddDigitalSignature(string _fileTempName, SPItemEventProperties properties)
        {

            object Visible = false;
            object readonlyfile = false;

            try
            {

                object missing = System.Reflection.Missing.Value;
                Microsoft.Office.Interop.Word.ApplicationClass wordapp = new Microsoft.Office.Interop.Word.ApplicationClass();
                Microsoft.Office.Interop.Word.Document wordDocument = wordapp.Documents.Open(ref
                    _fileTempName, ref missing,
                        ref readonlyfile, ref missing, ref missing,
                            ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing,
                                    ref Visible, ref missing, ref missing,
                                        ref missing, ref missing);

                SignatureSet _signatureSet = wordDocument.Signatures;
                Signature objSignature = _signatureSet.Add();
                
                if (objSignature == null)
                {
                    // DocumentNotSigned(properties);
                    return false;
                }

                else
                {
                    _signatureSet.Commit();
                    object saveChanges = true;
                    wordDocument.Close(ref saveChanges, ref missing, ref missing);
                    wordapp.Quit(ref missing, ref missing, ref missing);
                    return true;
                }
            }
            catch
            {
                return false;
            }

        }   // cierre AddDigitalSignature


        // ----------------------------------------------------------------------------------


        private void WriteFileAgain(SPItemEventProperties properties, string TemporaryFile)
        {

            SPFile currentFile = properties.ListItem.File;
            string TempFilePath = TemporaryFolder + TemporaryFile;
            FileStream st = new FileStream(TempFilePath, FileMode.Open);
            properties.ListItem.ParentList.ParentWeb.AllowUnsafeUpdates = true;
            currentFile.CheckOut();
            currentFile.SaveBinary(st);
            this.DisableEventFiring();
            currentFile.CheckIn(string.Empty);
            currentFile.Publish(string.Empty);
            currentFile.Approve(string.Empty);
            st.Close();
            properties.ListItem.ParentList.ParentWeb.AllowUnsafeUpdates = false;
            this.EnableEventFiring();
            FileInfo deletedfile = new FileInfo(TemporaryFolder + TemporaryFile);
            deletedfile.Delete();

        }   // cierre WriteFileAgain


        // ----------------------------------------------------------------------------------


        private void DocumentNotSigned(SPItemEventProperties properties)
        {

            properties.ListItem.ParentList.ParentWeb.AllowUnsafeUpdates = true;
            SPFile currentFile = properties.ListItem.File;
            if (currentFile.CheckOutStatus == SPFile.SPCheckOutStatus.None)
            {
                currentFile.CheckOut();
            }

            this.DisableEventFiring();
            currentFile.CheckIn(string.Empty);
            currentFile.Publish(string.Empty);
            currentFile.Deny(string.Empty);
            properties.ListItem.ParentList.ParentWeb.AllowUnsafeUpdates = false;
            this.EnableEventFiring();

        }   // cierre DocumentNotSigned


        // ----------------------------------------------------------------------------------


    }
}
