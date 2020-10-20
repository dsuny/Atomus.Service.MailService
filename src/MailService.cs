using Atomus.Diagnostics;
using Atomus.Net;
using System;
using System.Collections.Generic;
using System.Data;
using System.Net.Mail;

namespace Atomus.Service
{
    public class MailService : IService
    {
        Response IService.Request(ServiceDataSet serviceDataSet)
        {
            //IService service;

            try
            {
                if (!serviceDataSet.ServiceName.Equals("Atomus.Service.MailService"))
                    throw new Exception("Not Atomus.Service.MailService");

                ((IServiceDataSet)serviceDataSet).CreateServiceDataTable();

                return this.Excute(serviceDataSet);
            }
            catch (AtomusException exception)
            {
                DiagnosticsTool.MyTrace(exception);
                return (Response)Factory.CreateInstance("Atomus.Service.Response", false, true, exception);
            }
            catch (Exception exception)
            {
                DiagnosticsTool.MyTrace(exception);
                return (Response)Factory.CreateInstance("Atomus.Service.Response", false, true, exception);
            }
        }

        private Response Excute(ServiceDataSet serviceDataSet)
        {
            IResponse response;
            IMail mail;
            MailMessage message;
            List<IAttachmentBytes> attachmentBytes;
            IAttachmentBytes attachment;

            try
            {
                mail = ((IMail)this.CreateInstance("Email"));
                //mail = new Mail();

                foreach (DataTable table in ((IServiceDataSet)serviceDataSet).DataTables)
                {
                    attachmentBytes = new List<IAttachmentBytes>();

                    foreach (DataRow dataRow in table.Rows)
                        if (dataRow["Bytes"] != null)
                        {
                            attachment = new AttachmentBytes();
                            attachment.Bytes = (byte[])dataRow["Bytes"];
                            attachment.FileName = (string)dataRow["FileName"];
                            attachment.MediaType = (string)dataRow["MediaType"];

                            attachmentBytes.Add(attachment);
                        }

                    message = mail.CreateMailMessage(
                        ((IServiceDataSet)serviceDataSet)[table.TableName].GetAttribute("From").ToString()
                        , ((IServiceDataSet)serviceDataSet)[table.TableName].GetAttribute("To").ToString()
                        , ((IServiceDataSet)serviceDataSet)[table.TableName].GetAttribute("CC").ToString()
                        , ((IServiceDataSet)serviceDataSet)[table.TableName].GetAttribute("BCC").ToString()
                        , ((IServiceDataSet)serviceDataSet)[table.TableName].GetAttribute("ReplyTo").ToString()
                        , ((IServiceDataSet)serviceDataSet)[table.TableName].GetAttribute("Subject").ToString()
                        , ((IServiceDataSet)serviceDataSet)[table.TableName].GetAttribute("Body").ToString()
                        , this.EncodingConvert(((IServiceDataSet)serviceDataSet)[table.TableName].GetAttribute("Encoding").ToString())
                        , ((IServiceDataSet)serviceDataSet)[table.TableName].GetAttribute("IsBodyHtml").ToString().ToBool()
                        , attachmentBytes
                        , (MailPriority)Enum.Parse(typeof(MailPriority), ((IServiceDataSet)serviceDataSet)[table.TableName].GetAttribute("MailPriority").ToString())
                        , (DeliveryNotificationOptions)Enum.Parse(typeof(DeliveryNotificationOptions), ((IServiceDataSet)serviceDataSet)[table.TableName].GetAttribute("DeliveryNotificationOptions").ToString())
                        );
                    //message = mail.CreateMailMessage("dsun <dsun@dsun.kr>", "abc <kwonedea@daum.net>;abc2 <dsun@hanmail.net>", subject, body, System.Text.Encoding.UTF8, false, fileName);


                    mail.Send(((IServiceDataSet)serviceDataSet)[table.TableName].ConnectionName, message);
                }

                response = (IResponse)Factory.CreateInstance("Atomus.Service.Response", false, true);
                response.Status = Status.OK;

                return (Response)response;
            }
            catch (AtomusException exception)
            {
                DiagnosticsTool.MyTrace(exception);
                return (Response)Factory.CreateInstance("Atomus.Service.Response", false, true, exception);
            }
            catch (Exception exception)
            {
                DiagnosticsTool.MyTrace(exception);
                return (Response)Factory.CreateInstance("Atomus.Service.Response", false, true, exception);
            }
        }


        System.Text.Encoding EncodingConvert(string encoding)
        {
            switch (encoding)
            {
                case "us-ascii":
                    return System.Text.Encoding.ASCII;

                case "utf-16":
                    return System.Text.Encoding.Unicode;

                case "utf-32":
                    return System.Text.Encoding.UTF32;

                case "utf-7":
                    return System.Text.Encoding.UTF7;

                case "utf-8":
                    return System.Text.Encoding.UTF8;

                default:
                    throw new AtomusException("default type Not Support.");
            }
        }
    }
}
