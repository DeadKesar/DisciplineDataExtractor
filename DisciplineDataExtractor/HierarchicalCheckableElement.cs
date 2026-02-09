//using JetBrains.Annotations;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;

namespace DisciplineDataExtractor
{
    public abstract class HierarchicalCheckableElement : INotifyPropertyChanged
    {
        private bool isChecked;

        public HierarchicalCheckableElement Parent { get; set; }

        public IEnumerable<HierarchicalCheckableElement> Nodes => GetNodes();

        protected abstract IEnumerable<HierarchicalCheckableElement> GetNodes();

        public bool IsChecked
        {
            get => isChecked;
            set
            {
                foreach (var node in Nodes)
                    node.IsChecked = value;
                isChecked = value;

                var parent = Parent;
                while (parent != null)
                {
                    if (parent.Nodes.All(node => !node.IsChecked))
                    {
                        parent.isChecked = false;
                        parent.OnPropertyChanged(nameof(IsChecked));
                        parent = parent.Parent;
                    }
                    else
                    {
                        parent.isChecked = true;
                        parent.OnPropertyChanged(nameof(IsChecked));
                        parent = parent.Parent;
                    }
                }

                OnPropertyChanged(nameof(IsChecked));
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        //[NotifyPropertyChangedInvocator]
        public void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
