using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Threading.Tasks;

public class CodeSnippet
{
	public CodeSnippet()
	{
        [HttpPost]
        [Route("api/Project/GroupRegistrationBulkUpload")]
        [SwaggerResponse(HttpStatusCode.BadRequest, "groupId and groupUniqueCode is required \r\n\n"
       + "No file uploaded.\r\n\n"
       + "Only Excel files (.xlsx) are allowed.\r\n\n"
       + "No worksheets found in the Excel file.\r\n\n"
       + "Data is more than the total open slots \r\n\n"
       + "Invalid header spellings or Invalid header sequence.\r\n\n")]
        [SwaggerResponse(HttpStatusCode.NotFound, "The group was not found or was deleted")]
        [SwaggerResponse(HttpStatusCode.InternalServerError, EErrorCodeMessages.Error)]
        [SwaggerResponse(HttpStatusCode.UnsupportedMediaType, "Media type not supported.")]
        [SwaggerResponse((HttpStatusCode)422, "Project is cancelled or finished")]
        [SwaggerResponse((HttpStatusCode)460, "Spots does not available\r\n\n")]
        [SwaggerResponse(HttpStatusCode.OK, "File uploaded and processed successfully. \r\n\n" +
        "No data for uploading. \r\n\n")]
        public async Task<HttpResponseMessage> GroupRegistrationBulkUpload()
        {
            try
            {
                MobileServiceContext context = new MobileServiceContext();
                HttpRequest request = HttpContext.Current.Request;


                string groupId = request.Form["groupId"];
                string groupUniqueCode = request.Form["groupUniqueCode"];
                if (string.IsNullOrEmpty(groupUniqueCode) || string.IsNullOrEmpty(groupId))
                {
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "groupId and groupUniqueCode is required");
                }

                var getPartnerShiftGroupSet = await context.PartnerShiftGroupSet.Include(x => x.User).Where(x => x.Id == groupId
                    && x.GroupUniqueCode == groupUniqueCode).FirstOrDefaultAsync();
                if (getPartnerShiftGroupSet == null)
                {
                    return NotFoundResponse();
                }

                var projectData = await context.ProjectSet.Where(x => x.Id == getPartnerShiftGroupSet.ProjectId && !x.Deleted).FirstOrDefaultAsync();
                if (projectData.State == EProjectState.Canceled || projectData.UTCFinishedAt.AddMinutes(15) < DateTimeOffset.UtcNow)
                {
                    return Request.CreateResponse((HttpStatusCode)422);
                }

                bool AllowEventSpotToGroup = false; // If group spot is not available then before 30 min to start event to end event we allow .
                if (getPartnerShiftGroupSet.GroupOpenUserSlots <= 0)
                {
                    if (projectData.UTCStartedAt < DateTimeOffset.UtcNow.AddMinutes(30) && DateTimeUtils.AddMinutes(projectData.UTCFinishedAt) > DateTimeOffset.UtcNow)
                    {
                        if (projectData.OpenUserSlots <= 0)
                        {
                            return Request.CreateResponse((HttpStatusCode)460);
                        }
                        AllowEventSpotToGroup = true;
                    }
                    else
                    {
                        return Request.CreateResponse((HttpStatusCode)460);
                    }
                }

                // Ensure the request is a multipart request
                if (!Request.Content.IsMimeMultipartContent())
                {
                    return Request.CreateErrorResponse(HttpStatusCode.UnsupportedMediaType, "Media type not supported.");
                }

                var provider = new MultipartMemoryStreamProvider();
                await Request.Content.ReadAsMultipartAsync(provider);

                // Check if any files were uploaded
                if (!provider.Contents.Any())
                {
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "No file uploaded.");
                }

                // Initialize variables to store the file name and bytes
                string fileName = null;
                byte[] fileBytes = null;

                // Process the uploaded files
                foreach (var fileContent in provider.Contents)
                {
                    // Check if this is a file content (not form data)
                    if (!string.IsNullOrEmpty(fileContent.Headers.ContentDisposition.FileName))
                    {
                        // Get the file name from the content disposition header
                        fileName = fileContent.Headers.ContentDisposition.FileName.Trim('"');

                        // Read the file content as bytes
                        fileBytes = await fileContent.ReadAsByteArrayAsync();

                        // You can choose to break here if you expect only one file
                        break;
                    }
                }

                // Check if a file was uploaded
                if (fileBytes == null || string.IsNullOrEmpty(fileName))
                {
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "No file uploaded.");
                }

                // Check if the uploaded file is an Excel file
                if (System.IO.Path.GetExtension(fileName) != ".xlsx")
                {
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Only Excel files (.xlsx) are allowed.");
                }

                using (var stream = new MemoryStream(fileBytes))
                using (var workbook = new XLWorkbook(stream))
                {
                    // Access the first worksheet
                    var worksheet = workbook.Worksheet(1);
                    int rowCount = worksheet.LastRowUsed().RowNumber();
                    var OpenSlots = getPartnerShiftGroupSet.GroupOpenUserSlots; // Group Open slots
                    if (AllowEventSpotToGroup)
                    {
                        OpenSlots = projectData.OpenUserSlots; // Project Open slots
                    }

                    if ((rowCount - 1) <= OpenSlots) // Verify that the number of open slots is greater or equal to the number of data being uploaded
                    {
                        int i = 0;

                        List<string> validHeaders = new List<string> { "FirstName", "LastName", "Email", "IsMinor" };
                        // Spell check all four headers individually
                        if (!AreHeadersValid(worksheet, validHeaders))
                        {
                            return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Invalid header spellings or Invalid header sequence.");
                        }
                        // Check header sequence
                        if (worksheet.Cell(1, 1).Value.ToString() != "FirstName" || worksheet.Cell(1, 2).Value.ToString() != "LastName" ||
                            worksheet.Cell(1, 3).Value.ToString() != "Email" || worksheet.Cell(1, 4).Value.ToString() != "IsMinor")
                        {
                            return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Invalid header sequence.");
                        }
                        var StrDateTime = projectData.StartedAt.ToString("dddd, MMM dd");
                        var StartTime = projectData.StartedAt.ToString("hh:mm tt").Replace('.', ':');
                        var manager = await context.ManagerSet.FirstOrDefaultAsync(r => r.Id == projectData.ManagerId);
                        var orgName = await context.AssociationSet.FirstOrDefaultAsync(r => !r.Deleted && r.Id == manager.AssociationId);
                       

                        List<Personalizations> listPersonalizations = new List<Personalizations>();
                        for (int row = 2; row <= rowCount; row++) // Assuming headers are in the first row
                        {
                            var firstName = worksheet.Cell(row, 1).Value.ToString();
                            var lastName = worksheet.Cell(row, 2).Value.ToString();
                            var email = worksheet.Cell(row, 3).Value.ToString();
                            var isMinorText = worksheet.Cell(row, 4).Value.ToString().Trim().ToLower();
                            bool isMinor = false;

                            if (isMinorText == "true")
                            {
                                isMinor = true;
                            }
                            var InvitationData = new PartnerGroupInvitees();
                            var RegisteredData = new ProjectSupportUser();
                            if (!string.IsNullOrEmpty(firstName) && !string.IsNullOrEmpty(lastName)) // First name and Last name required for save
                            {
                                bool isSendEmail = false;
                                if (!string.IsNullOrEmpty(email) && IsValidEmail(email))  // check if the data coming has email also it is valid, then check for duplication
                                {
                                    InvitationData = await context.PartnerGroupInviteesSet.Where(x => x.GroupId == groupId && !x.Deleted && x.Name.Trim().ToLower() == firstName.Trim().ToLower() && x.LastName.Trim().ToLower() == lastName.Trim().ToLower() && x.Email == email).FirstOrDefaultAsync();
                                    // to check if already registered with the same data ksw-2718
                                    RegisteredData = await context.ProjectSupportUserSet.Include(x => x.User).Where(x => x.PartnerShiftGroupId == groupId && !x.Deleted && x.User.FirstName.Trim().ToLower() == firstName.Trim().ToLower() && x.User.LastName.Trim().ToLower() == lastName.Trim().ToLower() && (x.User.AuthId == email || x.User.AuthProviderEmail == email)).FirstOrDefaultAsync();
                                    isSendEmail = true;
                                }
                                else
                                {
                                    email = "";
                                }
                                if ((InvitationData == null || InvitationData.Id == null) && (RegisteredData == null || RegisteredData.Id == null))
                                {
                                    var inviteGroup = new PartnerGroupInvitees()
                                    {
                                        Id = Guid.NewGuid().ToString("N"),
                                        Name = firstName,
                                        Email = email,
                                        PhoneNumber = null,
                                        GroupId = groupId,
                                        LastName = lastName,
                                        IsMinor = isMinor
                                    };
                                    inviteGroup = context.PartnerGroupInviteesSet.Add(inviteGroup);
                                    if (context.SaveChanges() > 0)
                                    {
                                        
                                        if (isSendEmail)
                                        {
                                            ToEmail bulkEmail = new ToEmail();
                                            Personalizations personalizations = new Personalizations();
                                            bulkEmail.email = email;
                                            bulkEmail.name = inviteGroup.Name;
                                            EmailToken token = new EmailToken
                                            {
                                                Username = !string.IsNullOrWhiteSpace(inviteGroup.Name) ? "Hi " + inviteGroup.Name + "," : "Hello,",
                                                Date = StrDateTime,
                                                Time = StartTime,
                                                ProjectName = projectData.Name,
                                                WebSiteUrl = WebSiteUrl,
                                                OrganizationName = orgName != null ? orgName.Name : string.Empty,
                                                ProjectAddress = projectData.Address,
                                                GroupLeaderName = getPartnerShiftGroupSet.User.FirstName,
                                                GroupName = getPartnerShiftGroupSet.PartnerGroupName,
                                                Task = projectData.Work
                                            };
                                            List<ToEmail> listBulkEmail = new List<ToEmail>();
                                            listBulkEmail.Add(bulkEmail);
                                            personalizations.to = listBulkEmail;
                                            personalizations.dynamicTemplateData = token;
                                            listPersonalizations.Add(personalizations);
                                        }
                                        var eventmanager = await context.ManagerSet.Include(m => m.Association).Where(m => m.Id == projectData.ManagerId).FirstOrDefaultAsync();
                                        if (eventmanager.Association.IsPartner && eventmanager.Association.Id == ConfigSettings.FMSCAssociation)
                                        {
                                            PartnerGroupInviteesWebhookDetails partnergroupinviteeswebhookdetails = new PartnerGroupInviteesWebhookDetails();
                                            partnergroupinviteeswebhookdetails.firstName = firstName;
                                            partnergroupinviteeswebhookdetails.lastName = lastName;
                                            partnergroupinviteeswebhookdetails.email = email;
                                            partnergroupinviteeswebhookdetails.isMinor = isMinor;    // Minor flag not updating
                                            partnergroupinviteeswebhookdetails.kindlyPartnerGroupId = groupId;
                                            partnergroupinviteeswebhookdetails.kindlyInvitationId = inviteGroup.Id;
                                            partnergroupinviteeswebhookdetails.kindlyShiftId = projectData.Id;
                                            partnergroupinviteeswebhookdetails.method = EWebhookCalls.AddInviteeInGroup;
                                            await InviteeWebhookCall(partnergroupinviteeswebhookdetails);


                                        }
                                        i++;
                                    }
                                }
                            }
                        }

                        if (i > 0) // check for if any invitee data stored or not then only we calculate the spots
                        {
                            getPartnerShiftGroupSet.KindlyInviteeUserCount = getPartnerShiftGroupSet.KindlyInviteeUserCount + i;
                            if (AllowEventSpotToGroup)
                            {
                                var prj = context.ProjectSet.Where(p => p.Id == projectData.Id).FirstOrDefault();
                                prj.RegisteredUserCount = prj.RegisteredUserCount + i;
                                prj.OpenUserSlots = prj.OpenUserSlots - i;
                                context.Entry(prj).State = EntityState.Modified;
                                await context.SaveChangesAsync();
                            }
                            else
                            {
                                getPartnerShiftGroupSet.GroupOpenUserSlots = getPartnerShiftGroupSet.GroupOpenUserSlots - i;
                                context.Entry(getPartnerShiftGroupSet).State = EntityState.Modified;
                                await context.SaveChangesAsync();
                            }
                            // KSW-2815
                            var emailTemplateController = new EmailTemplateController
                            {
                                Request = new System.Net.Http.HttpRequestMessage(),
                                Configuration = new HttpConfiguration()
                            };
                            await emailTemplateController.AzureBusBulkEmailSend(listPersonalizations, (int)EEMailTemplateCode.GroupInvitation);
                        }

                        else
                        {
                            return Request.CreateResponse(HttpStatusCode.OK, "No data for uploading.");
                        }
                        return Request.CreateResponse(HttpStatusCode.OK, "File uploaded and processed successfully.");
                    }
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Data is more than the total open slots");
                }
            }
            catch (Exception ex)
            {
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, ex.Message);
            }
        }
    }
}
