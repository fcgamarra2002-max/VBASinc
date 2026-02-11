using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using VBIDE = Microsoft.Vbe.Interop;

namespace VBASinc.UI
{
    public partial class SelectComponentsForm : Form
    {
        private List<VBIDE.VBComponent> _allComponents;
        private TreeView? _treeView;
        private Button? _btnOK;
        private Button? _btnCancel;

        public List<VBIDE.VBComponent> SelectedComponents { get; private set; } = new List<VBIDE.VBComponent>();

        public SelectComponentsForm(List<VBIDE.VBComponent> components, string title = "Seleccionar Componentes")
        {
            _allComponents = components;
            InitializeComponent(title);
            PopulateTree();
        }

        private void InitializeComponent(string title)
        {
            this.Text = title;
            this.Size = new Size(500, 500);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.StartPosition = FormStartPosition.CenterParent;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            
            var label = new Label
            {
                Text = "Seleccione los componentes que desea exportar:",
                Location = new Point(15, 15),
                Size = new Size(450, 20),
                Font = new Font("Segoe UI", 9F, FontStyle.Bold)
            };
            
            _treeView = new TreeView
            {
                Location = new Point(15, 40),
                Size = new Size(455, 360),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom,
                CheckBoxes = true,
                ShowLines = true,
                ShowPlusMinus = true,
                ShowRootLines = true
            };
            
            _treeView.AfterCheck += TreeView_AfterCheck;
            
            _btnOK = new Button
            {
                Text = "Aceptar",
                Location = new Point(310, 415),
                Size = new Size(80, 30),
                DialogResult = DialogResult.OK
            };
            
            _btnCancel = new Button
            {
                Text = "Cancelar",
                Location = new Point(400, 415),
                Size = new Size(80, 30),
                DialogResult = DialogResult.Cancel
            };
            
            this.Controls.Add(label);
            this.Controls.Add(_treeView);
            this.Controls.Add(_btnOK);
            this.Controls.Add(_btnCancel);
            
            this.AcceptButton = _btnOK;
            this.CancelButton = _btnCancel;
        }

        private void PopulateTree()
        {
            if (_treeView == null) return;

            // Project Root
            var projectNode = new TreeNode("VBAProject") { Checked = true };
            _treeView.Nodes.Add(projectNode);

            // Groups
            var docNode = new TreeNode("Microsoft Excel Objetos") { Checked = true };
            var formNode = new TreeNode("Formularios") { Checked = true };
            var modNode = new TreeNode("Módulos") { Checked = true };
            var classNode = new TreeNode("Módulos de clase") { Checked = true };

            foreach (var comp in _allComponents)
            {
                var node = new TreeNode(comp.Name) { Tag = comp, Checked = true };
                
                switch (comp.Type)
                {
                    case VBIDE.vbext_ComponentType.vbext_ct_Document:
                        docNode.Nodes.Add(node);
                        break;
                    case VBIDE.vbext_ComponentType.vbext_ct_MSForm:
                        formNode.Nodes.Add(node);
                        break;
                    case VBIDE.vbext_ComponentType.vbext_ct_StdModule:
                        modNode.Nodes.Add(node);
                        break;
                    case VBIDE.vbext_ComponentType.vbext_ct_ClassModule:
                        classNode.Nodes.Add(node);
                        break;
                }
            }

            if (docNode.Nodes.Count > 0) projectNode.Nodes.Add(docNode);
            if (formNode.Nodes.Count > 0) projectNode.Nodes.Add(formNode);
            if (modNode.Nodes.Count > 0) projectNode.Nodes.Add(modNode);
            if (classNode.Nodes.Count > 0) projectNode.Nodes.Add(classNode);

            projectNode.ExpandAll();
        }

        private void TreeView_AfterCheck(object? sender, TreeViewEventArgs e)
        {
            if (e.Action == TreeViewAction.Unknown) return;

            // Update children recursively
            UpdateChildrenCheckedState(e.Node, e.Node.Checked);
            
            // Update parents (if all children unchecked, uncheck parent, etc.)
            UpdateParentCheckedState(e.Node);
        }

        private void UpdateChildrenCheckedState(TreeNode node, bool isChecked)
        {
            foreach (TreeNode child in node.Nodes)
            {
                child.Checked = isChecked;
                UpdateChildrenCheckedState(child, isChecked);
            }
        }

        private void UpdateParentCheckedState(TreeNode node)
        {
            if (node.Parent != null)
            {
                bool anyChecked = false;

                foreach (TreeNode sibling in node.Parent.Nodes)
                {
                    if (sibling.Checked) 
                    {
                        anyChecked = true;
                        break;
                    }
                }

                // If this is a category node or project node, we update it based on children
                // This is a simple logic: if any child is checked, parent stays/becomes checked?
                // Actually, more common: if all children are unchecked, uncheck parent.
                // If any child is checked, check parent.
                
                // Switching parent state programmatically
                // We don't want to trigger children update here, but AfterCheck handles it with Action.Unknown logic.
                if (node.Parent.Checked != anyChecked)
                {
                    node.Parent.Checked = anyChecked;
                    UpdateParentCheckedState(node.Parent);
                }
            }
        }

        private void OnOK()
        {
            SelectedComponents.Clear();
            if (_treeView == null) return;

            CollectCheckedComponents(_treeView.Nodes);
        }

        private void CollectCheckedComponents(TreeNodeCollection nodes)
        {
            foreach (TreeNode node in nodes)
            {
                if (node.Tag is VBIDE.VBComponent comp && node.Checked)
                {
                    SelectedComponents.Add(comp);
                }
                CollectCheckedComponents(node.Nodes);
            }
        }

        protected override void OnClosed(EventArgs e)
        {
            if (this.DialogResult == DialogResult.OK)
            {
                OnOK();
            }
            base.OnClosed(e);
        }
    }
}