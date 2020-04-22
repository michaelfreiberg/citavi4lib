using System;
using System.Collections.Generic;
using System.ComponentModel;

using System.IO;
using System.Net;
using System.Windows.Forms;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Web;

using SwissAcademic.Citavi;
using SwissAcademic.Citavi.Metadata;
using SwissAcademic.Citavi.Shell;
using SwissAcademic.Collections;


// Implementation of macro editor is preliminary and experimental.
// The Citavi object model is subject to change in future version.

public static class CitaviMacro {
	public static void Main() {
		
		//if we need a ref to the active project
		if (Program.ActiveProjectShell == null) return;
		SwissAcademic.Citavi.Project activeProject = Program.ActiveProjectShell.Project;
		
		if (activeProject == null)
		{
			DebugMacro.WriteLine("No active project.");
			return;
		}
		
		/*if this macro should ALWAYS affect all titles in active project, choose first option
		  if this macro should affect just filtered rows if there is a filter applied and ALL if not, choose second option
		  if this macro should affect just selected rows, choose third option */
			
		//ProjectReferenceCollection references = Program.ActiveProjectShell.Project.References;
		//List<Reference> references = Program.ActiveProjectShell.PrimaryMainForm.GetFilteredReferences();
		List < Reference > references = Program.ActiveProjectShell.PrimaryMainForm.GetSelectedReferences();
		
		if (references == null || !references.Any())
		{
			DebugMacro.WriteLine("No references selected.");
			return;
		}

		int countNotProvisional = 0;
		int countProvisional = 0;
		int countReferences = 0;

		foreach(Reference reference in references) {
			countReferences = countReferences + 1; 	
			
			string budgetCode = "Budget?";
			string specialistName = "Michael Freiberg";
			string edition = "p";
			string price = "? €";
			string taskType = "?";
			string orderNote = "-";
			string normalizedISBN = reference.Isbn.Isbn13.ToString();
			string buchhandelURL = "https://www.buchhandel.de/jsonapi/products?filter[products][query]=(is=" + normalizedISBN + ")";
			
			// Get availability information of the title from "buchhandel.de"
			try {
				Cursor.Current = Cursors.WaitCursor;
				
				DebugMacro.WriteLine("Getting availability information from " + buchhandelURL);
				WebClient client = new WebClient();
				client.Headers["Accept"] = "application/vnd.api+json";
				var jsonString = client.DownloadString(buchhandelURL);
				// DebugMacro.WriteLine(jsonString);

							
				// Get budget code from groups			
				List<SwissAcademic.Citavi.Group> refGroups = reference.Groups.ToList();
				
				foreach (SwissAcademic.Citavi.Group refGroup in refGroups)
				{
					Match match = Regex.Match(refGroup.Name, @"budget\:\s*(.+)", RegexOptions.IgnoreCase);
					if (match.Success)
					{
					   budgetCode = match.Groups[1].Value;
					}
				}

				// Quick and dirty extraction of price information via RegEx.
				// Parsing of JSON response would be preferred, but I could not handle
				// the concepts of System.Runtime.Serialization and System.Runtime.Serialization.Json;
				
				
					
		        // Get first match. Book price in Austria.
				Match matchPrice = Regex.Match(jsonString, @"\""value\""\:(\d+\.\d+)");
				if (matchPrice.Success)
				{
					var priceAT = matchPrice.Groups[1].Value;
					// DebugMacro.WriteLine(priceAT);
				}

				// Get second match. Book price in Germany.
				matchPrice = matchPrice.NextMatch();
				if (matchPrice.Success)
				{
					var priceDE = matchPrice.Groups[1].Value;
					price = Regex.Replace(priceDE, @"\.", ",") + " €";
					// DebugMacro.WriteLine(priceDE);
				}
					
		        Match matchAvail = Regex.Match(jsonString, @"\""provisional\""\:true");
		        if (matchAvail.Success)
				{
					reference.Groups.Add("Noch nicht erschienen");
					countProvisional = countProvisional + 1;
				}
				else
				{
					countNotProvisional = countNotProvisional + 1;
					
					foreach (SwissAcademic.Citavi.Group refGroup in refGroups)
					{
						Match match = Regex.Match(refGroup.Name, @"Noch nicht erschienen", RegexOptions.IgnoreCase);
						if (match.Success)
						{
							reference.Groups.Remove(refGroup);
						}
					}
					
					taskType = "Bestellen";
					orderNote = "Verfügbar. " + edition + " " + price;
					
									
					TaskItem order = reference.Tasks.Add(taskType);

					// Assign specialist.
					ContactInfo taskAssignee = activeProject.Contacts.GetMembers().FirstOrDefault<ContactInfo>(item => item.FullName.Equals(specialistName));
					if (taskAssignee == null) 
					{
						DebugMacro.WriteLine("No such contact.");
						return;
					}
					order.AssignedTo = taskAssignee.Key;
					
					order.Notes = orderNote;
				}
			}
			catch(Exception e) {
				Cursor.Current = Cursors.Default;
				DebugMacro.WriteLine("Could not read file from " + buchhandelURL + ": " + e.Message);
				return;
			}
			
			


		}
		
		// Message upon completion
		string message = "{0} Titel wurden geprüft. {1} sind noch nicht erschienen. {2} sind lieferbar.";
		message = string.Format(message, countReferences.ToString(), countProvisional.ToString(), countNotProvisional.ToString());
		MessageBox.Show(message, "Bericht", MessageBoxButtons.OK, MessageBoxIcon.Information);
	}
}

