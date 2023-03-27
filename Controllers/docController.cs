using System;
using System.Collections.Generic;
using System.IO;

namespace ExcelLikeAPI
{
    /// <summary>
    /// Represents an Excel-like session, which can contain multiple workbooks.
    /// </summary>
    public class Session
    {
        /// <summary>
        /// Gets the list of open workbooks in this session.
        /// </summary>
        public List<Workbook> Workbooks { get; private set; }
        /// <summary>
        /// Initializes a new instance of the <see cref="Session"/> class.
        /// </summary>
        public Session()
        {
            Workbooks = new List<Workbook>();
        }

        /// <summary>
        /// Opens a workbook from the specified file path.
        /// </summary>
        /// <param name="path">The file path of the workbook to open.</param>
        /// <returns>The opened workbook, or null if the file was not found.</returns>
        public Workbook OpenWorkbook(string path)
        {
            if (!File.Exists(path))
            {
                Console.WriteLine($"Error: File '{path}' not found.");
                return null;
            }

            var workbook = new Workbook(this, path);
            Workbooks.Add(workbook);
            workbook.Open();

            return workbook;
        }

        /// <summary>
        /// Creates a new, empty workbook with the specified name.
        /// </summary>
        /// <param name="name">The name of the new workbook.</param>
        /// <returns>The created workbook.</returns>
        public Workbook CreateWorkbook(string name)
        {
            var workbook = new Workbook(this, name);
            Workbooks.Add(workbook);

            return workbook;
        }

        /// <summary>
        /// Closes the specified workbook.
        /// </summary>
        /// <param name="workbook">The workbook to close.</param>
        public void CloseWorkbook(Workbook workbook)
        {
            Workbooks.Remove(workbook);
            workbook.Close();
        }
    }

    /// <summary>
    /// Represents an Excel-like workbook, which can contain multiple worksheets.
    /// </summary>
    public class Workbook
    {
        private Session _session;

        /// <summary>
        /// Gets the list of worksheets in this workbook.
        /// </summary>
        public List<Worksheet> Worksheets { get; private set; }

        /// <summary>
        /// Gets or sets the name of this workbook.
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Gets or sets the file path of this workbook.
        /// </summary>
        public string Path { get; set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="Workbook"/> class from the specified file path.
        /// </summary>
        /// <param name="session">The session that owns this workbook.</param>
        /// <param name="path">The file path of the workbook.</param>
        public Workbook(Session session, string path)
        {
            _session = session;
            Path = path;
            Name = System.IO.Path.GetFileName(path);
            Worksheets = new List<Worksheet>();
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Workbook"/> class with the specified name.
        /// </summary>
        /// <param name="session">The session that owns this workbook.</param>
        /// <param name="name">The name of the workbook.</param>
        public Workbook(Session session, string name)
        {
            _session = session;
            Name = name;
            Worksheets = new List<Worksheet>();
        }

        /// <summary>
        /// Opens this workbook.
        /// </summary>
        public void Open()
        {
            Console.WriteLine($"Opening workbook '{Name}' fromfile '{Path}'...");
            // implementation to open the workbook from file
        }
        /// <summary>
        /// Saves this workbook.
        /// </summary>
        public void Save()
        {
            Console.WriteLine($"Saving workbook '{Name}' to file '{Path}'...");
            // implementation to save the workbook to file
        }

        /// <summary>
        /// Saves this workbook with a new file path.
        /// </summary>
        /// <param name="newPath">The new file path to save the workbook to.</param>
        public void SaveAs(string newPath)
        {
            Console.WriteLine($"Saving workbook '{Name}' to new file '{newPath}'...");
            // implementation to save the workbook to the new file path
            Path = newPath;
        }

        /// <summary>
        /// Closes this workbook.
        /// </summary>
        public void Close()
        {
            Console.WriteLine($"Closing workbook '{Name}'...");
            // implementation to close the workbook
        }

        /// <summary>
        /// Adds a new worksheet to this workbook with the specified name.
        /// </summary>
        /// <param name="name">The name of the new worksheet.</param>
        /// <returns>The added worksheet.</returns>
        public Worksheet AddWorksheet(string name)
        {
            var worksheet = new Worksheet(this, name);
            Worksheets.Add(worksheet);

            return worksheet;
        }

        /// <summary>
        /// Deletes the specified worksheet from this workbook.
        /// </summary>
        /// <param name="worksheet">The worksheet to delete.</param>
        public void DeleteWorksheet(Worksheet worksheet)
        {
            Worksheets.Remove(worksheet);
        }

        /// <summary>
        /// Renames the specified worksheet in this workbook.
        /// </summary>
        /// <param name="worksheet">The worksheet to rename.</param>
        /// <param name="newName">The new name of the worksheet.</param>
        public void RenameWorksheet(Worksheet worksheet, string newName)
        {
            worksheet.Name = newName;
        }

        /// <summary>
        /// Makes the specified worksheet in this workbook active.
        /// </summary>
        /// <param name="worksheet">The worksheet to make active.</param>
        public void ActivateWorksheet(Worksheet worksheet)
        {
            // implementation to activate the specified worksheet
        }
    }

    /// <summary>
    /// Represents an Excel-like worksheet, which contains cells and data.
    /// </summary>
    public class Worksheet
    {
        private Workbook _workbook;

        /// <summary>
        /// Gets or sets the name of this worksheet.
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="Worksheet"/> class with the specified name.
        /// </summary>
        /// <param name="workbook">The workbook that owns this worksheet.</param>
        /// <param name="name">The name of the worksheet.</param>
        public Worksheet(Workbook workbook, string name)
        {
            _workbook = workbook;
            Name = name;
        }
    }
}
