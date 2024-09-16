using System.Net.Mail;
using OfficeOpenXml;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System.IO;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Linq;
using System;
using ExcelUploadAPI.Model_Class;
using NETCore.MailKit.Core;
using System.Net;
using Microsoft.Extensions.Configuration;
using MimeKit;
using Microsoft.Extensions.Options;
using SendingEmail.Controllers;

[ApiController]
[Route("api/[controller]")]
public class ExcelUploadController : ControllerBase
{

    private readonly EmailService _emailService;
    private readonly IConfiguration _configuration;


    [HttpPost("UploadExcel")]
    public async Task<IActionResult> UploadExcel(IFormFile file)
    {
        if (file == null || file.Length == 0)
        {
            return BadRequest("No file uploaded.");
        }

        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        List<UserModel> users = new List<UserModel>();
        List<string> errors = new List<string>(); // To collect error messages

        using (var stream = new MemoryStream())
        {
            await file.CopyToAsync(stream);

            try
            {
                using (var package = new ExcelPackage(stream))
                {
                    var worksheet = package.Workbook.Worksheets.FirstOrDefault();

                    if (worksheet == null)
                    {
                        return BadRequest("The uploaded Excel file doesn't contain any worksheets.");
                    }

                    int rowCount = worksheet.Dimension?.Rows ?? 0;

                    if (rowCount == 0)
                    {
                        return BadRequest("The uploaded Excel file is empty.");
                    }

                    for (int row = 2; row <= rowCount; row++)  // Start from row 2 (skipping headers)
                    {
                        string cellValue = worksheet.Cells[row, 1].Text;
                        int id;
                        bool isValidId = int.TryParse(cellValue, out id);  // First column must be an integer

                        if (!isValidId)
                        {
                            // Collect errors with more details
                            errors.Add($"Invalid data format in row {row}. Expected an integer in the first column. Found: '{cellValue}'");
                            continue;  // Skip this row
                        }

                        string name = worksheet.Cells[row, 2].Text;
                        string email = worksheet.Cells[row, 3].Text;

                        users.Add(new UserModel
                        {
                            Id = id,
                            Name = name,
                            Email = email
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Internal server error: {ex.Message}");
            }
        }

        if (errors.Any())
        {
            return BadRequest(new
            {
                message = "Some rows contained invalid data.",
                validData = users,
                errors
            });
        }

        return Ok(users);
    }

    [HttpPost("send")]
    public async Task<IActionResult> SendEmail([FromBody] EmailRequest request)
    {
        if (string.IsNullOrEmpty(request.To) || string.IsNullOrEmpty(request.Subject) || string.IsNullOrEmpty(request.Body))
        {
            return BadRequest("Email, subject, and body are required.");
        }

        try
        {
            // SMTP client configuration
            using (var smtpClient = new SmtpClient("smtp.your-email-provider.com")
            {
                Port = 465, // Replace with your SMTP port number
                Credentials = new NetworkCredential("keerthisri1106@gmail.com", "AppKey"),
                EnableSsl = true,
            })
            {
                var mailMessage = new MailMessage
                {
                    From = new MailAddress("poojastee30@gmail.com"),
                    Subject = request.Subject,
                    Body = request.Body,
                    IsBodyHtml = true, // Set to true if the body is HTML
                };

                mailMessage.To.Add(request.To);

                // Send email asynchronously
                await smtpClient.SendMailAsync(mailMessage);
            }

            return Ok("Email sent successfully.");
        }
        catch (SmtpException ex)
        {
            // Handle SMTP exceptions (e.g., network issues, authentication errors)
            return StatusCode(500, $"SMTP error: {ex.Message}");
        }
        catch (System.Exception ex)
        {
            // Handle general exceptions
            return StatusCode(500, $"Error sending email: {ex.Message}");
        }
    }


    public class EmailRequest
    {
        public string To { get; set; }
        public string Subject { get; set; }
        public string Body { get; set; }
    }

}























