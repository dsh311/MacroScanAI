/*
 * Copyright (C) 2025 David S. Shelley <davidsmithshelley@gmail.com>
 *
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License 
 * along with this program. If not, see <https://www.gnu.org/licenses/>.
 */

using OpenMcdf;
using System.Windows.Media.Imaging;


namespace MacroScanAI.Controls.TreeAndEditor
{
    public static class OleTreeBuilder
    {
        public static OleNode BuildTree(RootStorage rootStorage)
        {
            string pathToFolderIcon = "Assets/folder.png";
            var rootNode = new OleNode
            {
                Name = "Root",
                IsStream = false,
                Root = rootStorage,
                Icon = new BitmapImage(new Uri(pathToFolderIcon, UriKind.Relative)),
                IsSelected = false,
                IsExpanded = false
            };

            PopulateChildren(rootStorage, rootNode);

            return rootNode;
        }

        private static void PopulateChildren(dynamic storage, OleNode parent)
        {
            foreach (var item in storage.EnumerateEntries())
            {
                if (item.Type == EntryType.Storage)
                {
                    // Open the nested storage
                    var nestedStorage = storage.OpenStorage(item.Name);

                    
                    string pathToFolderIcon = "Assets/folder.png";

                    // Create a node for the storage
                    var node = new OleNode
                    {
                        Name = item.Name,
                        IsStream = false,
                        Parent = parent,
                        Icon = new BitmapImage(new Uri(pathToFolderIcon, UriKind.Relative)),
                        IsSelected = false,
                        IsExpanded = false
                    };

                    // Recursively populate children
                    PopulateChildren(nestedStorage, node);

                    parent.Children.Add(node);
                }
                else if (item.Type == EntryType.Stream)
                {
                    // Open the actual stream to get Size
                    var stream = storage.OpenStream(item.Name);

                    string pathToFileIcon = "Assets/file.png";

                    parent.Children.Add(new OleNode
                    {
                        Name = item.Name,
                        IsStream = true,
                        Stream = stream,
                        Parent = parent,
                        Icon = new BitmapImage(new Uri(pathToFileIcon, UriKind.Relative)),
                        IsSelected = false,
                        IsExpanded = false
                    });
                }
            }
        }

    }


}
