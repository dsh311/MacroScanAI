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
using System.Windows.Media;
using System.ComponentModel;

namespace MacroScanAI.Controls.TreeAndEditor
{
    public class OleNode
    {
        public string Name { get; set; } = "";
        public bool IsStream { get; set; }

        public bool HasModuleLinage { get; set; } = false;
        public CfbStream? Stream { get; set; }
        public RootStorage? Root { get; set; }

        public OleNode? Parent { get; set; }
        public List<OleNode> Children { get; set; } = new List<OleNode>();

        public string DisplayName => IsStream && Stream != null
            ? $"{Name} ({Stream.Length} bytes)"
            : Name;

        private ImageSource? _icon;

        public ImageSource? Icon
        {
            get { return _icon; }
            set { _icon = value; }
        }

        private bool _isExpanded;
        public bool IsExpanded
        {
            get => _isExpanded;
            set { _isExpanded = value; OnPropertyChanged(nameof(IsExpanded)); }
        }

        private bool _isSelected;
        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                if (_isSelected != value)
                {
                    _isSelected = value;
                    OnPropertyChanged(nameof(IsSelected));
                }
            }
        }

        // INotifyPropertyChanged implementation
        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged(string propertyName) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));


        public OleNode CloneSelf(OleNode node)
        {
            if (node == null) { return null; }

            var clone = new OleNode
            {
                Name = node.Name,
                IsStream = node.IsStream,
                HasModuleLinage = node.HasModuleLinage,
                Stream = node.Stream,
                Root = node.Root,
                Parent = node.Parent, // optionally null if you want detached copy
                IsExpanded = node.IsExpanded,
                Icon = node.Icon,
                Children = node.Children?.Select(CloneSelf).ToList()
            };

            // Update Parent references of children to the cloned node
            if (clone.Children != null)
            {
                foreach (var child in clone.Children)
                {
                    child.Parent = clone;
                }
            }

            return clone;
        }


    }


}
