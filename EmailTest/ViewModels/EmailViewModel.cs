using System;
using System.Net.Mail;
using System.Net;
using System.IO;
using System.Linq;

public class EmailViewModel
{
    private readonly EmailData _emailData;
    private readonly string _csvFolderPath;

    public EmailViewModel(EmailData emailData, string csvFolderPath = @"C:\AMV\DLDataLogs")
    {
        _emailData = emailData ?? throw new ArgumentNullException(nameof(emailData));
        _csvFolderPath = csvFolderPath;
    }

    public void SendLatestCsvEmail()
    {
        try
        {
            // Ensure folder exists
            if (!Directory.Exists(_csvFolderPath))
            {
                Console.WriteLine($"CSV folder not found: {_csvFolderPath}");
                return;
            }

            // Find the most recent CSV file
            string latestCsvFile = Directory.GetFiles(_csvFolderPath, "*.csv")
                .Where(f => Path.GetFileName(f).ToLower().Contains("shift"))
                .OrderByDescending(f => File.GetLastWriteTime(f))
                .FirstOrDefault();

            if (latestCsvFile == null)
            {
                Console.WriteLine("No CSV file found in the specified folder.");
                return;
            }

            string fileName = Path.GetFileNameWithoutExtension(latestCsvFile);
            Console.WriteLine($"Found latest CSV file: {Path.GetFileName(latestCsvFile)}");

            // Parse filename to extract shift and date
            // Format: S2_Report_Shift_3_11-Jan-2026
            string shift = "Unknown";
            string date = DateTime.Now.ToString("dd-MMM-yyyy");
            
            ParseFileName(fileName, out shift, out date);

            // Set attachment path and email content
            _emailData.AttachmentPath = latestCsvFile;
            _emailData.Subject = _emailData.Subject ?? "AMV Shift Report";
            
            // Format body with date and shift if template contains placeholders
            if (_emailData.Body != null && _emailData.Body.Contains("{0}") && _emailData.Body.Contains("{1}"))
            {
                _emailData.Body = string.Format(_emailData.Body, date, shift);
            }
            else
            {
                _emailData.Body = _emailData.Body ?? $"Please find attached the Shift Report for {date} corresponding to Shift {shift}.";
            }

            // Send email
            SendEmail();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error sending email: {ex.Message}");
            Console.WriteLine($"Stack trace: {ex.StackTrace}");
        }
    }

    private void ParseFileName(string fileName, out string shift, out string date)
    {
        shift = "Unknown";
        date = DateTime.Now.ToString("dd-MMM-yyyy");

        try
        {
            // Format: S2_Report_Shift_3_11-Jan-2026
            // Split by underscore
            string[] parts = fileName.Split('_');
            
            if (parts.Length >= 5)
            {
                // Find "Shift" index
                int shiftIndex = -1;
                for (int i = 0; i < parts.Length; i++)
                {
                    if (parts[i].Equals("Shift", StringComparison.OrdinalIgnoreCase))
                    {
                        shiftIndex = i;
                        break;
                    }
                }

                if (shiftIndex >= 0 && shiftIndex + 1 < parts.Length)
                {
                    // Get shift number (next part after "Shift")
                    shift = parts[shiftIndex + 1];
                    
                    // Get date (last part)
                    if (shiftIndex + 2 < parts.Length)
                    {
                        date = parts[shiftIndex + 2];
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error parsing filename: {ex.Message}");
        }
    }

    public bool SendEmail()
    {
        try
        {
            // Create the mail message
            using (MailMessage mail = new MailMessage())
            {
                mail.From = new MailAddress(_emailData.FromEmail);
                
                // Add multiple recipients from ToEmails list
                if (_emailData.ToEmails != null && _emailData.ToEmails.Any())
                {
                    foreach (var toEmail in _emailData.ToEmails)
                    {
                        mail.To.Add(toEmail);
                    }
                }
                else
                {
                    Console.WriteLine("No recipients specified.");
                    return false;
                }
                
                // CC
                if (_emailData.CcEmails != null && _emailData.CcEmails.Any())
                {
                    foreach (var ccEmail in _emailData.CcEmails)
                    {
                        mail.CC.Add(ccEmail);
                    }
                }
                
                mail.Subject = _emailData.Subject;
                mail.Body = _emailData.Body;
                mail.IsBodyHtml = false;

                // Attach the file if it exists
                if (!string.IsNullOrEmpty(_emailData.AttachmentPath) && File.Exists(_emailData.AttachmentPath))
                {
                    Attachment attachment = new Attachment(_emailData.AttachmentPath);
                    mail.Attachments.Add(attachment);
                }
                else if (!string.IsNullOrEmpty(_emailData.AttachmentPath))
                {
                    Console.WriteLine("Attachment file not found.");
                    return false;
                }

                // Configure the SMTP client
                using (SmtpClient smtpClient = new SmtpClient("smtp.gmail.com", 587))
                {
                    smtpClient.EnableSsl = true;
                    smtpClient.UseDefaultCredentials = false;
                    smtpClient.Credentials = new NetworkCredential(_emailData.FromEmail, _emailData.AppPassword);
                    smtpClient.DeliveryMethod = SmtpDeliveryMethod.Network;

                    // Send the email
                    smtpClient.Send(mail);
                    Console.WriteLine("Email sent successfully!");
                    return true;
                }
            }
        }
        catch (SmtpException ex)
        {
            Console.WriteLine($"SMTP Error: {ex.Message}");
            Console.WriteLine($"Status Code: {ex.StatusCode}");
            if (ex.InnerException != null)
                Console.WriteLine($"Inner Exception: {ex.InnerException.Message}");
            return false;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"General Error: {ex.Message}");
            return false;
        }
    }
}