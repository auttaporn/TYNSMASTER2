using System;
using System.Web;
using System.Data;
using System.Web.UI;
using System.Configuration;
using System.Web.UI.WebControls;

namespace VTS.Web.UI
{
    public partial class ParentChildTreeView : System.Web.UI.UserControl
    {
        #region Property

        public DataTable DataSource
        {
            get;
            set;
        }

        public String DisplayMember
        {
            get;
            set;
        }

        public String ValueMember
        {
            get;
            set;
        }

        public String ParentMember
        {
            get;
            set;
        }
        public String NavigateMember
        {
            get;
            set;
        }
        public String KeyMember
        {
            get;
            set;
        }

        #endregion

        #region event

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!Page.IsPostBack)
            {
                CommonLoad();
            }
        }

        #endregion

        #region Method

        private void CommonLoad()
        {
            PopulateTree(treeView);
        }

        private void PopulateTree(TreeView objTreeView)
        {
            if (DataSource != null)
            {
                foreach (DataRow dataRow in DataSource.Rows)
                {
                    if (dataRow[ParentMember] == DBNull.Value)
                    {
                        TreeNode treeRoot = new TreeNode();
                        treeRoot.Text = dataRow[DisplayMember].ToString();
                        treeRoot.Value = dataRow[ValueMember].ToString();
                        treeRoot.ImageUrl = "~/Images/Folder.gif";
                        treeRoot.Target = dataRow[NavigateMember].ToString();
                        treeRoot.ExpandAll();
                        objTreeView.Nodes.Add(treeRoot);                                               

                        foreach (TreeNode childnode in GetChildNode(Convert.ToInt64(dataRow[KeyMember])))
                        {
                            treeRoot.ChildNodes.Add(childnode);
                        }
                    }
                }
            }
        }


        private TreeNodeCollection GetChildNode(long parentid)
        {
            TreeNodeCollection childtreenodes = new TreeNodeCollection();
            DataView dataView1 = new DataView(DataSource);
            String strFilter = "" + ParentMember + "=" + parentid.ToString() + "";
            dataView1.RowFilter = strFilter;

            if (dataView1.Count > 0)
            {
                foreach (DataRow dataRow in dataView1.ToTable().Rows)
                {
                    TreeNode childNode = new TreeNode();
                    childNode.Text = dataRow[DisplayMember].ToString();
                    childNode.Value = dataRow[ValueMember].ToString();
                    childNode.Target = dataRow[NavigateMember].ToString();
                    childNode.ImageUrl = "~/Images/oInboxF.gif";
                    childNode.ExpandAll();

                    foreach (TreeNode cnode in GetChildNode(Convert.ToInt64(dataRow[KeyMember])))
                    {
                        childNode.ChildNodes.Add(cnode);
                    }
                    childtreenodes.Add(childNode);
                }
            }
            return childtreenodes;
        }

        private void CheckChild()
        {
        }

        #endregion
        //protected  void treeView_SelectedNodeChanged(object sender, EventArgs e)
        //{
        //    //Response.Redirect(
        //}


       public event EventHandler buttonClick;

        protected void treeView_SelectedNodeChanged(object sender, EventArgs e)
        {
            buttonClick(sender, e);
        }

}
}
