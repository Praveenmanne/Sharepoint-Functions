using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Diagnostics;
using System.ServiceModel;
using System.Text;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Utilities;


namespace asdf
{
    class Program
    {
        static void Main(string[] args)
        {
            {
                try
                {
                    //Console.Clear();
                    //Console.WriteLine("Please input site url : ");

                    const string siteUrl = "http://sp-live:8000/";

                    using (var site = new SPSite(siteUrl))
                    {
                        using (var web = site.OpenWeb())
                        {

                            var headers = new StringDictionary();

                            headers.Add("from", "sender@domain.com");
                            headers.Add("to", "praveen@tillidsoft.com");
                            headers.Add("subject", "Welcome to the SharePoint");
                            headers.Add("fAppendHtmlTag", "True"); //To enable HTML Tags

                            var strMessage = new System.Text.StringBuilder();
                            strMessage.Append("Message from CEO:");

                            strMessage.Append("<span style='color:red;'> Make sure you have completed the survey! </span>");
                            SPUtility.SendEmail(web, headers, strMessage.ToString());

                            //SPListItemCollection employeetype = GetListItemCollection(web.Lists.TryGetList("Employee Leaves"), "Employee Type", "Probationary", "Leave Type", "Paid Leave");

                            //foreach (SPListItem currentUseremptypeDetail in employeetype)
                            //{
                            //    currentUseremptypeDetail["Leave Balance"] =
                            //        Convert.ToInt16(currentUseremptypeDetail["Leave Balance"]) + 1;
                            //    currentUseremptypeDetail.Update();
                            //}

                        }

                    }


                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error:");
                    Console.WriteLine(ex.Message);
                }
                finally
                {
                    Console.WriteLine("Process Completed");
                    Console.ReadKey();
                }
            }
        }

        internal SPListItemCollection GetListItemCollection(SPList spList, string key, string value)
        {
            // Return list item collection based on the lookup field

            SPField spField = spList.Fields[key];
            var query = new SPQuery
            {
                Query = @"<Where>
                        <Eq>
                            <FieldRef Name='" + spField.InternalName + @"'/><Value Type='" + spField.Type.ToString() + @"'>" + value + @"</Value>
                        </Eq>
                        </Where>"
            };

            return spList.GetItems(query);
        }

        internal SPListItemCollection GetListItemCollection(SPList spList, string keyOne, string valueOne, string keyTwo, string valueTwo)
        {
            // Return list item collection based on the lookup field

            SPField spFieldOne = spList.Fields[keyOne];
            SPField spFieldTwo = spList.Fields[keyTwo];
            var query = new SPQuery
            {
                Query = @"<Where>
                          <And>
                                <Eq>
                                    <FieldRef Name=" + spFieldOne.InternalName + @" />
                                    <Value Type=" + spFieldOne.Type.ToString() + ">" + valueOne + @"</Value>
                                </Eq>
                                <Eq>
                                    <FieldRef Name=" + spFieldTwo.InternalName + @" />
                                    <Value Type=" + spFieldTwo.Type.ToString() + ">" + valueTwo + @"</Value>
                                </Eq>
                          </And>
                        </Where>"
            };

            return spList.GetItems(query);
        }

        internal SPListItemCollection GetListItemCollection(SPList spList, string keyOne, string valueOne, string keyTwo, string valueTwo, string keyThree, string valueThree)
        {
            // Return list item collection based on the lookup field

            SPField spFieldOne = spList.Fields[keyOne];
            SPField spFieldTwo = spList.Fields[keyTwo];
            SPField spFieldThree = spList.Fields[keyThree];
            var query = new SPQuery
            {
                Query = @"<Where>
                          <And>
                             <And>
                                <Eq>
                                   <FieldRef Name=" + spFieldOne.InternalName + @" />
                                   <Value Type=" + spFieldOne.Type.ToString() + @">" + valueOne + @"</Value>
                                </Eq>
                                <Eq>
                                   <FieldRef Name=" + spFieldTwo.InternalName + @" />
                                   <Value Type=" + spFieldTwo.Type.ToString() + @">" + valueTwo + @"</Value>
                                </Eq>
                             </And>
                             <Eq>
                                <FieldRef Name=" + spFieldThree.InternalName + @" />
                                <Value Type=" + spFieldThree.Type.ToString() + @">" + valueThree + @"</Value>
                             </Eq>
                          </And>
                       </Where>"
            };

            return spList.GetItems(query);
        }

        internal SPListItemCollection GetListItemCollection(SPList spList, string keyOne, string valueOne, string keyTwo, string valueTwo, string keyThree, string valueThree, string keyFour, string valueFour)
        {
            // Return list item collection based on the lookup field

            SPField spFieldOne = spList.Fields[keyOne];
            SPField spFieldTwo = spList.Fields[keyTwo];
            SPField spFieldThree = spList.Fields[keyThree];
            SPField spFieldFour = spList.Fields[keyFour];
            var query = new SPQuery
            {
                Query = @"<Where>
                          <And>
                             <And>
                                <And>
                                <Eq>
                                   <FieldRef Name=" + spFieldOne.InternalName + @" />
                                   <Value Type=" + spFieldOne.Type.ToString() + @">" + valueOne + @"</Value>
                                </Eq>
                                <Eq>
                                   <FieldRef Name=" + spFieldTwo.InternalName + @" />
                                   <Value Type=" + spFieldTwo.Type.ToString() + @">" + valueTwo + @"</Value>
                                </Eq>
                             </And>
                             <Eq>
                                <FieldRef Name=" + spFieldThree.InternalName + @" />
                                <Value Type=" + spFieldThree.Type.ToString() + @">" + valueThree + @"</Value>
                             </Eq>
                          </And>
                             <Eq>
                                <FieldRef Name=" + spFieldFour.InternalName + @" />
                                <Value Type=" + spFieldFour.Type.ToString() + @">" + valueFour + @"</Value>
                             </Eq>
                          </And>
                       </Where>"
            };

            return spList.GetItems(query);
        }

        /// <summary>
        /// Gets the SPFolder from SharePoint
        /// </summary>
        /// <param name="urlString">The SharePoint URL</param>
        /// <param name="documentLibrary">The Document Library Folder to retrieve</param>
        /// <param name="contextUser">The context User name</param>
        /// <param name="spBasePermission"> </param>
        /// <param name="customError">Parameter for returning the exception</param>
        /// <returns>A SPFolder based off the parameters</returns>
        internal SPFolder GetFolder(string urlString, string documentLibrary, string contextUser, SPBasePermissions spBasePermission, out FaultException customError)
        {
            customError = null;
            SPFolder spFolder = null;
            SPUserToken spUserTokenPermission;
           
            SPSecurity.RunWithElevatedPrivileges(delegate()
            {
                using (var site = new SPSite(urlString))
                {
                    using (var spWeb = site.OpenWeb())
                    {
                        var spUser = spWeb.AllUsers[contextUser];

                        // Check if the user has permissions to access the Sharepoint
                        if (!spWeb.DoesUserHavePermissions(spUser.LoginName, SPBasePermissions.Open))
                        {
                            // throw new FaultException(Properties.Resources.NOT_AUTHORIZED);
                            //customErrorReturnType = new FaultException(Properties.Resources.NOT_AUTHORIZED);
                            //return null;
                        }

                        spUserTokenPermission = spUser.UserToken;
                    }
                }

                using (var site = new SPSite(urlString, spUserTokenPermission))
                {
                    using (SPWeb spWeb = site.OpenWeb())
                    {
                        // Check if the user has base permissions.
                        if (!spWeb.DoesUserHavePermissions(spBasePermission))
                        {
                            //throw new FaultException(Properties.Resources.NOT_AUTHORIZED);
                            //customErrorReturnType = new FaultException(Properties.Resources.NOT_AUTHORIZED);
                            //return null;
                        }
                        spFolder = spWeb.GetFolder(documentLibrary);

                        //If document library doesn't exists, then throw exception
                        if (!spFolder.Exists)
                        {
                            // throw new FaultException(String.Format(Properties.Resources.LIBRARY_MISSING, documentLibrary));
                            //customErrorReturnType = new FaultException(String.Format(Properties.Resources.LIBRARY_MISSING, documentLibrary));
                            //return null;
                        }
                    }
                }
            });
            customError = null;
            return spFolder;

        }

        /// <summary>
        /// Gets the SPListItemCollection from Document library
        /// </summary>
        /// <param name="spWeb">The spweb</param>
        /// <param name="documentLibrary">The document library</param>
        /// <param name="documentName">The document name</param>
        /// <returns>SPListItemCollection based off the parameter</returns>
        internal SPListItemCollection GetListItemCollection(SPWeb spWeb, string documentLibrary, string documentName)
        {
            // Return list item collection based on the document name
            var stringBuilder = new StringBuilder();
            stringBuilder.Append("<Where><Eq><FieldRef Name='FileLeafRef'/><Value Type='File'>" + documentName + "</Value></Eq></Where>");

            var spQuery = new SPQuery();
            spQuery.Query = stringBuilder.ToString();
            spQuery.RowLimit = 1;

            return spWeb.Lists[documentLibrary].GetItems(spQuery);
        }
        /// <summary>
        /// Validates the meta data.
        /// </summary>
        /// <param name="urlString">The URL string.</param>
        /// <param name="metaData">The meta data.</param>
        /// <param name="documentLibrary">The document library.</param>
        /// <param name="customError">Parameter for returning the exception.</param>
        internal void ValidateMetaData(string urlString, Dictionary<string, string> metaData, string documentLibrary, out FaultException customError)
        {
            customError = null;
            string columnNames = string.Empty;
            string requiredColumnNames = string.Empty;

            SPSecurity.RunWithElevatedPrivileges(delegate()
            {
                using (var site = new SPSite(urlString))
                {
                    using (var spWeb = site.OpenWeb())
                    {
                        SPFieldCollection spFields = spWeb.GetFolder(documentLibrary).DocumentLibrary.Fields;

                        // verify field is available in field collection
                        foreach (KeyValuePair<String, String> entry in metaData)
                        {
                            if (!spFields.ContainsField(entry.Key.ToString()))
                            {
                                columnNames += entry.Key.ToString() + ", ";
                            }
                            else
                            {
                                // verify require field have valid value.
                                SPField field = spFields.GetField(entry.Key.ToString());

                                if (field.Required && string.IsNullOrEmpty(entry.Value.ToString()))
                                {
                                    requiredColumnNames += entry.Key.ToString() + ", ";
                                }
                            }
                        }

                        if (!string.IsNullOrEmpty(requiredColumnNames))
                        {
                            //throw new FaultException(string.Format(Properties.Resources.REQUIRED_FIELDS_DOESNT_HAVE_VALUE, requiredColumnNames.Substring(0, requiredColumnNames.Length - 2), documentLibrary));
                        }

                        if (!string.IsNullOrEmpty(columnNames))
                        {
                            // throw new FaultException(string.Format(Properties.Resources.METADATA_NOT_AVAILABLE, columnNames.Substring(0, columnNames.Length - 2), documentLibrary));
                        }
                    }
                }
            });
        }

        /// <summary>
        /// Converts to dictionary.
        /// </summary>
        /// <param name="metaData">The meta data.</param>
        /// <param name="customError">Parameter for returning the exception.</param>
        /// <returns>dictionary</returns>
        internal Dictionary<string, string> ConvertToDictionary(string metaData, out FaultException customError)
        {
            customError = null;

            if (!string.IsNullOrEmpty(metaData))
            {
                var returnHashTable = new Dictionary<string, string>();
                try
                {
                    // spliting using the seperators
                    var metaDataSeparators = new string[] { "~|#" };
                    var split = metaData.Split(metaDataSeparators, StringSplitOptions.None);

                    foreach (string t in split)
                    {
                        var valueSeparators = new string[] { "~|~" };
                        var value = t.Split(valueSeparators, StringSplitOptions.None);
                        returnHashTable.Add(value[0].Trim(), value[1].Trim());
                    }
                }
                catch (Exception e)
                {
                    Trace.TraceError(DateTime.Now + ": " + e.Message.ToString() + Environment.NewLine + e.StackTrace);
                    customError = new FaultException(e.Message);
                }

                return returnHashTable;
            }
            else
            {
                //customError = new FaultException(string.Format(Properties.Resources.NOT_VALID_STRING, "MetaData"));
            }
            return null;
        }


        /// <summary>
        /// Gets a file from a SharePoint document library
        /// </summary>
        /// <param name="urlString">The complete SharePoint URL</param>
        /// <param name="documentLibrary">The name of the document library</param>
        /// <param name="documentName">The name of the file to get</param>
        /// <param name="contextUser">The context User name</param>
        /// <param name="spBasePermission">The base permission</param>
        /// <param name="customError">Parameter for returning exception, if any</param>
        /// <returns>A SPFile based off the parameters</returns>
        internal SPFile GetFile(string urlString, string documentLibrary, string documentName, string contextUser, SPBasePermissions spBasePermission, out FaultException customError)
        {
            customError = null;
            SPFile spFile = null;

            var folder = this.GetFolder(urlString, documentLibrary, contextUser, spBasePermission, out customError);

            if (customError != null) { throw customError; }

            //If document library doesn't exists, then throw exception
            if (!folder.Exists)
            {
                // customError = new FaultException(String.Format(Properties.Resources.LIBRARY_MISSING, documentLibrary));
                return null;
            }
            SPWeb spWeb;

            SPSecurity.RunWithElevatedPrivileges(delegate()
            {
                using (var site = new SPSite(urlString))
                {
                    using (spWeb = site.OpenWeb())
                    {
                        var spUserTokenPermission = spWeb.AllUsers[contextUser].UserToken;
                        using (var site1 = new SPSite(urlString, spUserTokenPermission))
                        {
                            using (spWeb = site1.OpenWeb())
                            {
                                // Get the file
                                spFile = spWeb.GetFile(urlString + "/" + documentLibrary + "/" + documentName);
                            }
                        }
                    }
                }
            });

            return spFile;
        }






    }
}