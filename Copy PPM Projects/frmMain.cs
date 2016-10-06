using Microsoft.ProjectServer.Client;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Net;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace Copy_PPM_Projects
{

    public partial class frmMain : System.Windows.Forms.Form
    {
        private static ProjectContext ppmContext;
        private static ProjectContext newppmContext;

        public frmMain()
        {
            InitializeComponent();
        }

        private void btnCarregar_Click(object sender, EventArgs e)
        {
            // Conectando no PPM Express
            ppmContext = new ProjectContext(txtPWAUrlOrigem.Text);
            
            var secpass = new SecureString();
            string pass = txtPwdPPM.Text;
            pass.ToList().ForEach(secpass.AppendChar);

            newppmContext.Credentials = new SharePointOnlineCredentials(txtUsuarioPPM.Text, secpass);

            ppmContext.Load(ppmContext.Projects);
            ppmContext.ExecuteQuery();

            // Conectando no novo ambiente
            newppmContext = new ProjectContext(txtPWAUrlDestino.Text);

            secpass = new SecureString();
            pass = txtSenha.Text;
            pass.ToList().ForEach(secpass.AppendChar);

            newppmContext.Credentials = new SharePointOnlineCredentials(txtUsuario.Text, secpass);

            // Carregando projetos
            newppmContext.Load(newppmContext.Projects);
            newppmContext.ExecuteQuery();

            // Carregando projetos
            newppmContext.Load(newppmContext.CustomFields);
            newppmContext.ExecuteQuery();

            progressBar.Maximum = ppmContext.Projects.Count();

            foreach (var prjOrigem in ppmContext.Projects)
            {
                bool projectReadyToCopy = false;
                try
                {
                    var prjOrigemComCustomFields = prjOrigem.IncludeCustomFields;
                    ppmContext.Load(prjOrigemComCustomFields);
                    ppmContext.ExecuteQuery();

                    var prjDestino = newppmContext.Projects.FirstOrDefault(c => c.Name == prjOrigemComCustomFields.Name);
                    if (prjDestino.IsCheckedOut)
                    {
                        var retry = MessageBox.Show("Project with check-out: " + prjDestino.Name + "\nYou must check-in the Project to copy the fields values. \nDo you want to try again?", "Check-out Project", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if(retry == DialogResult.Yes)
                        {
                            projectReadyToCopy = true;
                        }
                    }
                    else
                    {
                        projectReadyToCopy = true;
                    }

                    if(projectReadyToCopy)
                    {
                        var prjDestinoComCustomFields = prjDestino.CheckOut().IncludeCustomFields;
                        newppmContext.Load(prjDestinoComCustomFields);
                        newppmContext.ExecuteQuery();

                        foreach (var cfOrigem in prjOrigemComCustomFields.FieldValues)
                        {
                            var cfDetalhe = newppmContext.CustomFields.FirstOrDefault(c => c.InternalName == cfOrigem.Key);
                            if (cfDetalhe.Formula == null)
                            {
                                prjDestinoComCustomFields.SetCustomFieldValue(cfOrigem.Key, cfOrigem.Value);
                            }
                        }

                        prjDestinoComCustomFields.Owner = prjOrigemComCustomFields.Owner;

                        prjDestinoComCustomFields.Update();
                        var queueJob = prjDestinoComCustomFields.Publish(true);
                        newppmContext.WaitForQueue(queueJob, 20);

                        var pubProject = newppmContext.Projects.GetByGuid(prjDestinoComCustomFields.Id);
                        pubProject.SubmitToWorkflow();
                        newppmContext.ExecuteQuery();

                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }

                progressBar.Value += 1;

            }
        }
    }
}
