using System;
using SharpShell.Attributes;
using SharpShell.SharpContextMenu;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using FolderPowerTools.Properties;

namespace FolderPowerTools
{
    [ComVisible(true)]
    [COMServerAssociation(AssociationType.Class, @"Directory\Background")]
    public class CreateDateFolder : SharpContextMenu
    {
        protected override bool CanShowMenu()
        {
            return true;
        }

        protected override ContextMenuStrip CreateMenu()
        {
            //  Create the context menu.
            var contextMenu = new ContextMenuStrip();

            //  Add the 'copy path' item. This just copies the folder path for the extension.
            var menuItem = new ToolStripMenuItem
            {
                Text = @"Create Date Folder",
                Image = Properties.Resources.Add_Folder_icon
            };

            // Ovo kopira folder u Clipboard
            //menuItem.Click += (sender, args) => Clipboard.SetText(FolderPath);
            menuItem.Click += (sender, args) => CreateFolder();
            contextMenu.Items.Add(menuItem);

            //  Return the menu.
            return contextMenu;
        }
        private string GetCurrentDate()
        {
            var currentDate = DateTime.Today;
            return currentDate.ToString("yyyy-MM-dd");
        }
        private void CreateFolder()
        {
            var currentDate = GetCurrentDate();
            var currentPath = "";
            if (FolderPath != null)
            {
                currentPath = FolderPath;
            }

            var pathString = System.IO.Path.Combine(currentPath, currentDate);
            Clipboard.SetText(pathString);

            try
            {
                System.IO.Directory.CreateDirectory(pathString);
            }
            catch (Exception exception)
            {
                var builder = new StringBuilder();
                builder.AppendLine($"Greska prilikom kreiranja foldera {pathString}: {exception.Message}");
                MessageBox.Show(builder.ToString(), Resources.CreateDateFolder_CreateFolder_ERROR_, MessageBoxButtons.OK);
            }
        }
    }
}
