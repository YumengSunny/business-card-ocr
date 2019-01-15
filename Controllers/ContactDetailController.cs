using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Text;
using System.Threading.Tasks;
using BCardReader.Modals;
using BCardReader.Utility;
using Microsoft.AspNetCore.Mvc;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Blob;

namespace BCardReader.Controllers
{
    [Route("api/contactdetail")]
    [ApiController]
    public class ContactDetailController : Controller
    {
        public const string excelPath = "D:\\home\\excel\\google\\";

        public const string dataConnection = "Server=tcp:sqlserverocr.database.windows.net,1433;Initial Catalog = user_data; Persist Security Info=False;User ID = Developer1; Password=Canada123;MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout = 30;";

        public const string storageConnection = "DefaultEndpointsProtocol=https;AccountName=cardexcel;AccountKey=0R49QVKF4SXc7R0F9iXIKlU1F37+5BUOMX9evJKKOVyqRNZsGa9TyWwbB6zd/2hXu9Fz9ZXYRnLdhhMQoY5jBQ==;EndpointSuffix=core.windows.net";

        public const string azureServerPath = "https://cardexcel.blob.core.windows.net/excelfilecontainer/";




        // GET api/contactdetail
        [HttpGet]
        public List<ContactDetail> GetAllContactDetail()
        {
            List<ContactDetail> contactDetails = new List<ContactDetail>();

            SqlConnection con = new SqlConnection(dataConnection);
            try
            {

                StringBuilder queryString = new StringBuilder("select CONTACT_ID,NAME,TITLE,COMPANY_NAME,TEL,MOBILE_NO,EMAIL,WEBSITE," +
                       " REMARK1,REMARK2,REMARK3,REMARK4,REMARK5,REMARK6,REMARK7,REMARK8,REMARK9,REMARK10,ENT_DATE,MODIFY_DATE " +
                       " from contact_detail where 1=1");

                SqlCommand cmd = new SqlCommand(queryString.ToString(), con);
                con.Open();

                SqlDataReader rdr = cmd.ExecuteReader();

                while (rdr.Read())
                {
                    ContactDetail contactDetail = new ContactDetail();

                    contactDetail.ContactId = rdr["CONTACT_ID"].ToString();
                    contactDetail.Name = rdr["NAME"].ToString().Trim();
                    contactDetail.Title = rdr["TITLE"].ToString().Trim();
                    contactDetail.CompanyName = rdr["COMPANY_NAME"].ToString().Trim();
                    contactDetail.Tel = rdr["TEL"].ToString().Trim();
                    contactDetail.Mobile = rdr["MOBILE_NO"].ToString();
                    contactDetail.Email = rdr["EMAIL"].ToString().Trim();
                    contactDetail.Website = rdr["WEBSITE"].ToString().Trim();
                    contactDetail.Remark1 = rdr["REMARK1"].ToString().Trim();
                    contactDetail.Remark2 = rdr["REMARK2"].ToString().Trim();
                    contactDetail.Remark3 = rdr["REMARK3"].ToString().Trim();
                    contactDetail.Remark4 = rdr["REMARK4"].ToString().Trim();
                    contactDetail.Remark5 = rdr["REMARK5"].ToString().Trim();
                    contactDetail.Remark6 = rdr["REMARK6"].ToString().Trim();
                    contactDetail.Remark7 = rdr["REMARK7"].ToString().Trim();
                    contactDetail.Remark8 = rdr["REMARK8"].ToString().Trim();
                    contactDetail.Remark9 = rdr["REMARK9"].ToString().Trim();
                    contactDetail.Remark10 = rdr["REMARK10"].ToString().Trim();

                    contactDetails.Add(contactDetail);
                }


                con.Close();
            }
            catch (Exception e)
            {
                return contactDetails;
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                }
            }

            return contactDetails;

        }

        // GET api/contactdetail
        [HttpPost("retrieve")]
        public List<ContactDetail> GetContactDetail([FromBody] ContactDetail contactDetailForm)
        {
            List<ContactDetail> contactDetails = new List<ContactDetail>();

            SqlConnection con = new SqlConnection(dataConnection);
            try
            {

                StringBuilder queryString= new StringBuilder("select CONTACT_ID,NAME,TITLE,COMPANY_NAME,TEL,MOBILE_NO,EMAIL,WEBSITE," +
                       " REMARK1,REMARK2,REMARK3,REMARK4,REMARK5,REMARK6,REMARK7,REMARK8,REMARK9,REMARK10,ENT_DATE,MODIFY_DATE " +
                       " from contact_detail where 1=1");
                
                if(contactDetailForm.ContactId!=null && !contactDetailForm.ContactId.Equals(""))
                {
                    queryString.Append(" and CONTACT_ID=@CONTACT_ID ");
                }
                if (contactDetailForm.Name != null && !contactDetailForm.Name.Equals(""))
                {
                    queryString.Append(" and NAME=@NAME ");
                }
                if (contactDetailForm.Title != null && !contactDetailForm.Title.Equals(""))
                {
                    queryString.Append(" and TITLE=@TITLE ");
                }
                if (contactDetailForm.CompanyName != null && !contactDetailForm.CompanyName.Equals(""))
                {
                    queryString.Append(" and COMPANY_NAME=@COMPANY_NAME ");
                }
                if (contactDetailForm.Tel != null && !contactDetailForm.Tel.Equals(""))
                {
                    queryString.Append(" and TEL=@TEL ");
                }
                if (contactDetailForm.Mobile != null && !contactDetailForm.Mobile.Equals(""))
                {
                    queryString.Append(" and MOBILE_NO=@MOBILE_NO ");
                }
                if (contactDetailForm.Email != null && !contactDetailForm.Email.Equals(""))
                {
                    queryString.Append(" and EMAIL=@EMAIL ");
                }
                if (contactDetailForm.Website != null && !contactDetailForm.Website.Equals(""))
                {
                    queryString.Append(" and WEBSITE=@WEBSITE ");
                }
                if (contactDetailForm.Remark1 != null && !contactDetailForm.Remark1.Equals(""))
                {
                    queryString.Append(" and REMARK1=@REMARK1 ");
                }
                if (contactDetailForm.Remark2 != null && !contactDetailForm.Remark2.Equals(""))
                {
                    queryString.Append(" and REMARK2=@REMARK2 ");
                }
                if (contactDetailForm.Remark3 != null && !contactDetailForm.Remark3.Equals(""))
                {
                    queryString.Append(" and REMARK3=@REMARK3 ");
                }
                if (contactDetailForm.Remark4 != null && !contactDetailForm.Remark4.Equals(""))
                {
                    queryString.Append(" and REMARK4=@REMARK4 ");
                }
                if (contactDetailForm.Remark5 != null && !contactDetailForm.Remark5.Equals(""))
                {
                    queryString.Append(" and REMARK5=@REMARK5 ");
                }
                if (contactDetailForm.Remark6 != null && !contactDetailForm.Remark6.Equals(""))
                {
                    queryString.Append(" and REMARK6=@REMARK6 ");
                }
                if (contactDetailForm.Remark7 != null && !contactDetailForm.Remark7.Equals(""))
                {
                    queryString.Append(" and REMARK7=@REMARK7 ");
                }
                if (contactDetailForm.Remark8 != null && !contactDetailForm.Remark8.Equals(""))
                {
                    queryString.Append(" and REMARK8=@REMARK8 ");
                }
                if (contactDetailForm.Remark9 != null && !contactDetailForm.Remark9.Equals(""))
                {
                    queryString.Append(" and REMARK9=@REMARK9 ");
                }
                if (contactDetailForm.Remark10 != null && !contactDetailForm.Remark10.Equals(""))
                {
                    queryString.Append(" and REMARK10=@REMARK10 ");
                }


                SqlCommand cmd = new SqlCommand(queryString.ToString(), con);

                cmd.Parameters.AddWithValue("@CONTACT_ID", contactDetailForm.ContactId != null ? contactDetailForm.ContactId.Trim() : "");
                cmd.Parameters.AddWithValue("@NAME", contactDetailForm.Name != null ? contactDetailForm.Name.Trim() : "");
                cmd.Parameters.AddWithValue("@TITLE", contactDetailForm.Title != null ? contactDetailForm.Title.Trim() : "");
                cmd.Parameters.AddWithValue("@COMPANY_NAME", contactDetailForm.CompanyName != null ? contactDetailForm.CompanyName.Trim() : "");
                cmd.Parameters.AddWithValue("@TEL", contactDetailForm.Tel != null ? contactDetailForm.Tel.Trim() : "");
                cmd.Parameters.AddWithValue("@MOBILE_NO", contactDetailForm.Mobile != null ? contactDetailForm.Mobile.Trim() : "");
                cmd.Parameters.AddWithValue("@EMAIL", contactDetailForm.Email != null ? contactDetailForm.Email.Trim() : "");
                cmd.Parameters.AddWithValue("@WEBSITE", contactDetailForm.Website != null ? contactDetailForm.Website.Trim() : "");
                cmd.Parameters.AddWithValue("@REMARK1", contactDetailForm.Remark1 != null ? contactDetailForm.Remark1.Trim() : "");
                cmd.Parameters.AddWithValue("@REMARK2", contactDetailForm.Remark2 != null ? contactDetailForm.Remark2.Trim() : "");
                cmd.Parameters.AddWithValue("@REMARK3", contactDetailForm.Remark3 != null ? contactDetailForm.Remark3.Trim() : "");
                cmd.Parameters.AddWithValue("@REMARK4", contactDetailForm.Remark4 != null ? contactDetailForm.Remark4.Trim() : "");
                cmd.Parameters.AddWithValue("@REMARK5", contactDetailForm.Remark5 != null ? contactDetailForm.Remark5.Trim() : "");
                cmd.Parameters.AddWithValue("@REMARK6", contactDetailForm.Remark6 != null ? contactDetailForm.Remark6.Trim() : "");
                cmd.Parameters.AddWithValue("@REMARK7", contactDetailForm.Remark7 != null ? contactDetailForm.Remark7.Trim() : "");
                cmd.Parameters.AddWithValue("@REMARK8", contactDetailForm.Remark8 != null ? contactDetailForm.Remark8.Trim() : "");
                cmd.Parameters.AddWithValue("@REMARK9", contactDetailForm.Remark9 != null ? contactDetailForm.Remark9.Trim() : "");
                cmd.Parameters.AddWithValue("@REMARK10", contactDetailForm.Remark10 != null ? contactDetailForm.Remark10.Trim() : "");

                con.Open();

                SqlDataReader rdr = cmd.ExecuteReader();

                while (rdr.Read())
                {
                    ContactDetail contactDetail = new ContactDetail();

                    contactDetail.ContactId = rdr["CONTACT_ID"].ToString();
                    contactDetail.Name = rdr["NAME"].ToString().Trim();
                    contactDetail.Title = rdr["TITLE"].ToString().Trim();
                    contactDetail.CompanyName= rdr["COMPANY_NAME"].ToString().Trim();
                    contactDetail.Tel= rdr["TEL"].ToString().Trim();
                    contactDetail.Mobile = rdr["MOBILE_NO"].ToString();
                    contactDetail.Email = rdr["EMAIL"].ToString().Trim();
                    contactDetail.Website = rdr["WEBSITE"].ToString().Trim();
                    contactDetail.Remark1 = rdr["REMARK1"].ToString().Trim();
                    contactDetail.Remark2 = rdr["REMARK2"].ToString().Trim();
                    contactDetail.Remark3 = rdr["REMARK3"].ToString().Trim();
                    contactDetail.Remark4 = rdr["REMARK4"].ToString().Trim();
                    contactDetail.Remark5 = rdr["REMARK5"].ToString().Trim();
                    contactDetail.Remark6 = rdr["REMARK6"].ToString().Trim();
                    contactDetail.Remark7 = rdr["REMARK7"].ToString().Trim();
                    contactDetail.Remark8 = rdr["REMARK8"].ToString().Trim();
                    contactDetail.Remark9 = rdr["REMARK9"].ToString().Trim();
                    contactDetail.Remark10 = rdr["REMARK10"].ToString().Trim();

                    contactDetails.Add(contactDetail);
                }
              

                con.Close();
            }
            catch (Exception e)
            {
              return contactDetails;
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                }
            }

            return contactDetails;
        }

        // POST api/contactdetail/add
        [HttpPost("add")]
        public String AddContactDetail([FromBody] ContactDetail contactDetail)
        {
            SqlConnection con = new SqlConnection(dataConnection);
            try
            {
                StringBuilder stringBuilder = new StringBuilder("insert into contact_detail(NAME,TITLE,COMPANY_NAME,TEL,MOBILE_NO,EMAIL,WEBSITE," +
                       " REMARK1,REMARK2,REMARK3,REMARK4,REMARK5,REMARK6,REMARK7,REMARK8,REMARK9,REMARK10,ENT_DATE,MODIFY_DATE)" +
                       " values(@NAME,@TITLE,@COMPANY_NAME,@TEL,@MOBILE_NO,@EMAIL,@WEBSITE,@REMARK1,@REMARK2,@REMARK3,@REMARK4," +
                       "       @REMARK5,@REMARK6,@REMARK7,@REMARK8,@REMARK9,@REMARK10,@ENT_DATE,@MODIFY_DATE)");
              
                SqlCommand cmd = new SqlCommand(stringBuilder.ToString(), con);

                con.Open();

                cmd.Parameters.AddWithValue("@NAME", contactDetail.Name !=null ? contactDetail.Name.Trim() : "");
                cmd.Parameters.AddWithValue("@TITLE", contactDetail.Title !=null ? contactDetail.Title.Trim(): "");
                cmd.Parameters.AddWithValue("@COMPANY_NAME", contactDetail.CompanyName != null ? contactDetail.CompanyName.Trim(): "");
                cmd.Parameters.AddWithValue("@TEL", contactDetail.Tel != null ? contactDetail.Tel .Trim(): "");
                cmd.Parameters.AddWithValue("@MOBILE_NO", contactDetail.Mobile != null ? contactDetail.Mobile.Trim() : "");
                cmd.Parameters.AddWithValue("@EMAIL", contactDetail.Email != null ? contactDetail.Email.Trim(): "");
                cmd.Parameters.AddWithValue("@WEBSITE", contactDetail.Website != null ? contactDetail.Website.Trim() : "");
                cmd.Parameters.AddWithValue("@REMARK1", contactDetail.Remark1 != null ? contactDetail.Remark1.Trim() : "");
                cmd.Parameters.AddWithValue("@REMARK2", contactDetail.Remark2 != null ? contactDetail.Remark2.Trim() : "");
                cmd.Parameters.AddWithValue("@REMARK3", contactDetail.Remark3 != null ? contactDetail.Remark3.Trim() : "");
                cmd.Parameters.AddWithValue("@REMARK4", contactDetail.Remark4 != null ? contactDetail.Remark4.Trim() : "");
                cmd.Parameters.AddWithValue("@REMARK5", contactDetail.Remark5 != null ? contactDetail.Remark5.Trim() : "");
                cmd.Parameters.AddWithValue("@REMARK6", contactDetail.Remark6 != null ? contactDetail.Remark6.Trim() : "");
                cmd.Parameters.AddWithValue("@REMARK7", contactDetail.Remark7 != null ? contactDetail.Remark7.Trim() : "");
                cmd.Parameters.AddWithValue("@REMARK8", contactDetail.Remark8 != null ? contactDetail.Remark8.Trim() : "");
                cmd.Parameters.AddWithValue("@REMARK9", contactDetail.Remark9 != null ? contactDetail.Remark9.Trim() : "");
                cmd.Parameters.AddWithValue("@REMARK10", contactDetail.Remark10 != null ? contactDetail.Remark10.Trim() : "");
                cmd.Parameters.AddWithValue("@ENT_DATE", DateTime.Now);
                cmd.Parameters.AddWithValue("@MODIFY_DATE", DateTime.Now);


                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch(Exception e)
            {
                return "Operation Failed";
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                }
            }

            return "Operation Success";
        }

        // PUT api/contactdetail/update/5
        [HttpPost("update/{id}")]
        public String UpdateContactDetail(int id, [FromBody] ContactDetail contactDetail)
        {
            SqlConnection con = new SqlConnection(dataConnection);
            try
            {

                StringBuilder stringBuilder = new StringBuilder("update contact_detail set MODIFY_DATE=@MODIFY_DATE ");

           
                if (contactDetail.Name != null)
                {
                    stringBuilder.Append(" , NAME=@NAME ");
                }
                if (contactDetail.Title != null)
                {
                    stringBuilder.Append(" , TITLE=@TITLE ");
                }
                if (contactDetail.CompanyName != null)
                {
                    stringBuilder.Append(" , COMPANY_NAME=@COMPANY_NAME ");
                }
                if (contactDetail.Tel != null)
                {
                    stringBuilder.Append(" , TEL=@TEL ");
                }
                if (contactDetail.Mobile != null)
                {
                    stringBuilder.Append(" , MOBILE_NO=@MOBILE_NO ");
                }
                if (contactDetail.Email != null)
                {
                    stringBuilder.Append(" , EMAIL=@EMAIL ");
                }
                if (contactDetail.Website != null)
                {
                    stringBuilder.Append(" , WEBSITE=@WEBSITE ");
                }
                if (contactDetail.Remark1 != null)
                {
                    stringBuilder.Append(" , REMARK1=@REMARK1 ");
                }
                if (contactDetail.Remark2 != null)
                {
                    stringBuilder.Append(" , REMARK2=@REMARK2 ");
                }
                if (contactDetail.Remark3 != null)
                {
                    stringBuilder.Append(" , REMARK3=@REMARK3 ");
                }
                if (contactDetail.Remark4 != null)
                {
                    stringBuilder.Append(" , REMARK4=@REMARK4 ");
                }
                if (contactDetail.Remark5 != null)
                {
                    stringBuilder.Append(" , REMARK5=@REMARK5 ");
                }
                if (contactDetail.Remark6 != null)
                {
                    stringBuilder.Append(" , REMARK6=@REMARK6 ");
                }
                if (contactDetail.Remark7 != null)
                {
                    stringBuilder.Append(" , REMARK7=@REMARK7 ");
                }
                if (contactDetail.Remark8 != null)
                {
                    stringBuilder.Append(" , REMARK8=@REMARK8 ");
                }
                if (contactDetail.Remark9 != null)
                {
                    stringBuilder.Append(" , REMARK9=@REMARK9 ");
                }
                if (contactDetail.Remark10 != null)
                {
                    stringBuilder.Append(" , REMARK10=@REMARK10 ");
                }

                stringBuilder.Append(" WHERE CONTACT_ID=@CONTACT_ID");

                SqlCommand cmd = new SqlCommand(stringBuilder.ToString(), con);

                con.Open();

                cmd.Parameters.AddWithValue("@NAME", contactDetail.Name != null ? contactDetail.Name.Trim() : "");
                cmd.Parameters.AddWithValue("@TITLE", contactDetail.Title != null ? contactDetail.Title.Trim() : "");
                cmd.Parameters.AddWithValue("@COMPANY_NAME", contactDetail.CompanyName != null ? contactDetail.CompanyName.Trim() : "");
                cmd.Parameters.AddWithValue("@TEL", contactDetail.Tel != null ? contactDetail.Tel.Trim() : "");
                cmd.Parameters.AddWithValue("@MOBILE_NO", contactDetail.Mobile != null ? contactDetail.Mobile.Trim() : "");
                cmd.Parameters.AddWithValue("@EMAIL", contactDetail.Email != null ? contactDetail.Email.Trim() : "");
                cmd.Parameters.AddWithValue("@WEBSITE", contactDetail.Website != null ? contactDetail.Website.Trim() : "");
                cmd.Parameters.AddWithValue("@REMARK1", contactDetail.Remark1 != null ? contactDetail.Remark1.Trim() : "");
                cmd.Parameters.AddWithValue("@REMARK2", contactDetail.Remark2 != null ? contactDetail.Remark2.Trim() : "");
                cmd.Parameters.AddWithValue("@REMARK3", contactDetail.Remark3 != null ? contactDetail.Remark3.Trim() : "");
                cmd.Parameters.AddWithValue("@REMARK4", contactDetail.Remark4 != null ? contactDetail.Remark4.Trim() : "");
                cmd.Parameters.AddWithValue("@REMARK5", contactDetail.Remark5 != null ? contactDetail.Remark5.Trim() : "");
                cmd.Parameters.AddWithValue("@REMARK6", contactDetail.Remark6 != null ? contactDetail.Remark6.Trim() : "");
                cmd.Parameters.AddWithValue("@REMARK7", contactDetail.Remark7 != null ? contactDetail.Remark7.Trim() : "");
                cmd.Parameters.AddWithValue("@REMARK8", contactDetail.Remark8 != null ? contactDetail.Remark8.Trim() : "");
                cmd.Parameters.AddWithValue("@REMARK9", contactDetail.Remark9 != null ? contactDetail.Remark9.Trim() : "");
                cmd.Parameters.AddWithValue("@REMARK10", contactDetail.Remark10 != null ? contactDetail.Remark10.Trim() : "");
                cmd.Parameters.AddWithValue("@CONTACT_ID",id.ToString());
                cmd.Parameters.AddWithValue("@MODIFY_DATE", DateTime.Now);

                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception e)
            {
                
                return "Operation Failed"+e.ToString();
            }
            finally
            {
                if(con!=null)
                {
                    con.Close();
                }
            }
            return "Operation Success";
        }

        // DELETE api/contactdetail/delete/
        [HttpPost("delete")]
        public List<ContactDetail> DeleteContactDetail(ContactDetail contactDetailForm)
        {

            List<ContactDetail> contactDetails = new List<ContactDetail>();
                SqlConnection con = new SqlConnection(dataConnection);
            try
            {

                StringBuilder queryString = new StringBuilder("select CONTACT_ID,NAME,TITLE,COMPANY_NAME,TEL,MOBILE_NO,EMAIL,WEBSITE," +
                      " REMARK1,REMARK2,REMARK3,REMARK4,REMARK5,REMARK6,REMARK7,REMARK8,REMARK9,REMARK10,ENT_DATE,MODIFY_DATE " +
                      " from contact_detail where 1=1 ");

                if(!((contactDetailForm.ContactId != null && !contactDetailForm.ContactId.Equals(""))
                    || (contactDetailForm.Name != null && !contactDetailForm.Name.Equals(""))
                    || (contactDetailForm.Title != null && !contactDetailForm.Title.Equals(""))
                    || (contactDetailForm.CompanyName != null && !contactDetailForm.CompanyName.Equals(""))
                    || (contactDetailForm.Tel != null && !contactDetailForm.Tel.Equals(""))
                    || (contactDetailForm.Mobile != null && !contactDetailForm.Mobile.Equals(""))
                    || (contactDetailForm.Email != null && !contactDetailForm.Email.Equals(""))
                    || (contactDetailForm.Website != null && !contactDetailForm.Website.Equals(""))
                    || (contactDetailForm.Remark1 != null && !contactDetailForm.Remark1.Equals(""))
                    || (contactDetailForm.Remark2 != null && !contactDetailForm.Remark2.Equals(""))
                    || (contactDetailForm.Remark3 != null && !contactDetailForm.Remark3.Equals(""))
                    || (contactDetailForm.Remark4 != null && !contactDetailForm.Remark4.Equals(""))
                    || (contactDetailForm.Remark5 != null && !contactDetailForm.Remark5.Equals(""))
                    || (contactDetailForm.Remark6 != null && !contactDetailForm.Remark6.Equals(""))
                    || (contactDetailForm.Remark7 != null && !contactDetailForm.Remark7.Equals(""))
                    || (contactDetailForm.Remark8 != null && !contactDetailForm.Remark8.Equals(""))
                    || (contactDetailForm.Remark9 != null && !contactDetailForm.Remark9.Equals(""))
                    || (contactDetailForm.Remark10 != null && !contactDetailForm.Remark10.Equals(""))
                   
                    ))
                {
                    queryString.Append(" and 1=2 ");
                }

                if (contactDetailForm.ContactId != null && !contactDetailForm.ContactId.Equals(""))
                {
                    queryString.Append(" and CONTACT_ID=@CONTACT_ID ");
                }
                if (contactDetailForm.Name != null && !contactDetailForm.Name.Equals(""))
                {
                    queryString.Append(" and NAME=@NAME ");
                }
                if (contactDetailForm.Title != null && !contactDetailForm.Title.Equals(""))
                {
                    queryString.Append(" and TITLE=@TITLE ");
                }
                if (contactDetailForm.CompanyName != null && !contactDetailForm.CompanyName.Equals(""))
                {
                    queryString.Append(" and COMPANY_NAME=@COMPANY_NAME ");
                }
                if (contactDetailForm.Tel != null && !contactDetailForm.Tel.Equals(""))
                {
                    queryString.Append(" and TEL=@TEL ");
                }
                if (contactDetailForm.Mobile != null && !contactDetailForm.Mobile.Equals(""))
                {
                    queryString.Append(" and MOBILE_NO=@MOBILE_NO ");
                }
                if (contactDetailForm.Email != null && !contactDetailForm.Email.Equals(""))
                {
                    queryString.Append(" and EMAIL=@EMAIL ");
                }
                if (contactDetailForm.Website != null && !contactDetailForm.Website.Equals(""))
                {
                    queryString.Append(" and WEBSITE=@WEBSITE ");
                }
                if (contactDetailForm.Remark1 != null && !contactDetailForm.Remark1.Equals(""))
                {
                    queryString.Append(" and REMARK1=@REMARK1 ");
                }
                if (contactDetailForm.Remark2 != null && !contactDetailForm.Remark2.Equals(""))
                {
                    queryString.Append(" and REMARK2=@REMARK2 ");
                }
                if (contactDetailForm.Remark3 != null && !contactDetailForm.Remark3.Equals(""))
                {
                    queryString.Append(" and REMARK3=@REMARK3 ");
                }
                if (contactDetailForm.Remark4 != null && !contactDetailForm.Remark4.Equals(""))
                {
                    queryString.Append(" and REMARK4=@REMARK4 ");
                }
                if (contactDetailForm.Remark5 != null && !contactDetailForm.Remark5.Equals(""))
                {
                    queryString.Append(" and REMARK5=@REMARK5 ");
                }
                if (contactDetailForm.Remark6 != null && !contactDetailForm.Remark6.Equals(""))
                {
                    queryString.Append(" and REMARK6=@REMARK6 ");
                }
                if (contactDetailForm.Remark7 != null && !contactDetailForm.Remark7.Equals(""))
                {
                    queryString.Append(" and REMARK7=@REMARK7 ");
                }
                if (contactDetailForm.Remark8 != null && !contactDetailForm.Remark8.Equals(""))
                {
                    queryString.Append(" and REMARK8=@REMARK8 ");
                }
                if (contactDetailForm.Remark9 != null && !contactDetailForm.Remark9.Equals(""))
                {
                    queryString.Append(" and REMARK9=@REMARK9 ");
                }
                if (contactDetailForm.Remark10 != null && !contactDetailForm.Remark10.Equals(""))
                {
                    queryString.Append(" and REMARK10=@REMARK10 ");
                }

                SqlCommand cmd = new SqlCommand(queryString.ToString(), con);

                cmd.Parameters.AddWithValue("@NAME", contactDetailForm.Name != null ? contactDetailForm.Name.Trim() : "");
                cmd.Parameters.AddWithValue("@TITLE", contactDetailForm.Title != null ? contactDetailForm.Title.Trim() : "");
                cmd.Parameters.AddWithValue("@COMPANY_NAME", contactDetailForm.CompanyName != null ? contactDetailForm.CompanyName.Trim() : "");
                cmd.Parameters.AddWithValue("@TEL", contactDetailForm.Tel != null ? contactDetailForm.Tel.Trim() : "");
                cmd.Parameters.AddWithValue("@MOBILE_NO", contactDetailForm.Mobile != null ? contactDetailForm.Mobile.Trim() : "");
                cmd.Parameters.AddWithValue("@EMAIL", contactDetailForm.Email != null ? contactDetailForm.Email.Trim() : "");
                cmd.Parameters.AddWithValue("@WEBSITE", contactDetailForm.Website != null ? contactDetailForm.Website.Trim() : "");
                cmd.Parameters.AddWithValue("@REMARK1", contactDetailForm.Remark1 != null ? contactDetailForm.Remark1.Trim() : "");
                cmd.Parameters.AddWithValue("@REMARK2", contactDetailForm.Remark2 != null ? contactDetailForm.Remark2.Trim() : "");
                cmd.Parameters.AddWithValue("@REMARK3", contactDetailForm.Remark3 != null ? contactDetailForm.Remark3.Trim() : "");
                cmd.Parameters.AddWithValue("@REMARK4", contactDetailForm.Remark4 != null ? contactDetailForm.Remark4.Trim() : "");
                cmd.Parameters.AddWithValue("@REMARK5", contactDetailForm.Remark5 != null ? contactDetailForm.Remark5.Trim() : "");
                cmd.Parameters.AddWithValue("@REMARK6", contactDetailForm.Remark6 != null ? contactDetailForm.Remark6.Trim() : "");
                cmd.Parameters.AddWithValue("@REMARK7", contactDetailForm.Remark7 != null ? contactDetailForm.Remark7.Trim() : "");
                cmd.Parameters.AddWithValue("@REMARK8", contactDetailForm.Remark8 != null ? contactDetailForm.Remark8.Trim() : "");
                cmd.Parameters.AddWithValue("@REMARK9", contactDetailForm.Remark9 != null ? contactDetailForm.Remark9.Trim() : "");
                cmd.Parameters.AddWithValue("@REMARK10", contactDetailForm.Remark10 != null ? contactDetailForm.Remark10.Trim() : "");
                cmd.Parameters.AddWithValue("@CONTACT_ID", contactDetailForm.ContactId != null ? contactDetailForm.ContactId.Trim() : "");

                con.Open();

                 

                   SqlDataReader rdr = cmd.ExecuteReader();

                   while (rdr.Read())
                   {
                       ContactDetail contactDetail = new ContactDetail();

                       contactDetail.ContactId = rdr["CONTACT_ID"].ToString();
                       contactDetail.Name = rdr["NAME"].ToString().Trim();
                       contactDetail.Title = rdr["TITLE"].ToString().Trim();
                       contactDetail.CompanyName = rdr["COMPANY_NAME"].ToString().Trim();
                       contactDetail.Tel = rdr["TEL"].ToString().Trim();
                       contactDetail.Mobile = rdr["MOBILE_NO"].ToString();
                       contactDetail.Email = rdr["EMAIL"].ToString().Trim();
                       contactDetail.Website = rdr["WEBSITE"].ToString().Trim();
                       contactDetail.Remark1 = rdr["REMARK1"].ToString().Trim();
                       contactDetail.Remark2 = rdr["REMARK2"].ToString().Trim();
                       contactDetail.Remark3 = rdr["REMARK3"].ToString().Trim();
                       contactDetail.Remark4 = rdr["REMARK4"].ToString().Trim();
                       contactDetail.Remark5 = rdr["REMARK5"].ToString().Trim();
                       contactDetail.Remark6 = rdr["REMARK6"].ToString().Trim();
                       contactDetail.Remark7 = rdr["REMARK7"].ToString().Trim();
                       contactDetail.Remark8 = rdr["REMARK8"].ToString().Trim();
                       contactDetail.Remark9 = rdr["REMARK9"].ToString().Trim();
                       contactDetail.Remark10 = rdr["REMARK10"].ToString().Trim();

                       contactDetails.Add(contactDetail);
                   }

                 con.Close();

                // Delete if any condition is there in input parameter

                if (((contactDetailForm.ContactId != null && !contactDetailForm.ContactId.Equals(""))
                  || (contactDetailForm.Name != null && !contactDetailForm.Name.Equals(""))
                  || (contactDetailForm.Title != null && !contactDetailForm.Title.Equals(""))
                  || (contactDetailForm.CompanyName != null && !contactDetailForm.CompanyName.Equals(""))
                  || (contactDetailForm.Tel != null && !contactDetailForm.Tel.Equals(""))
                  || (contactDetailForm.Mobile != null && !contactDetailForm.Mobile.Equals(""))
                  || (contactDetailForm.Email != null && !contactDetailForm.Email.Equals(""))
                  || (contactDetailForm.Website != null && !contactDetailForm.Website.Equals(""))
                  || (contactDetailForm.Remark1 != null && !contactDetailForm.Remark1.Equals(""))
                  || (contactDetailForm.Remark2 != null && !contactDetailForm.Remark2.Equals(""))
                  || (contactDetailForm.Remark3 != null && !contactDetailForm.Remark3.Equals(""))
                  || (contactDetailForm.Remark4 != null && !contactDetailForm.Remark4.Equals(""))
                  || (contactDetailForm.Remark5 != null && !contactDetailForm.Remark5.Equals(""))
                  || (contactDetailForm.Remark6 != null && !contactDetailForm.Remark6.Equals(""))
                  || (contactDetailForm.Remark7 != null && !contactDetailForm.Remark7.Equals(""))
                  || (contactDetailForm.Remark8 != null && !contactDetailForm.Remark8.Equals(""))
                  || (contactDetailForm.Remark9 != null && !contactDetailForm.Remark9.Equals(""))
                  || (contactDetailForm.Remark10 != null && !contactDetailForm.Remark10.Equals(""))

                  ))
                {

                    StringBuilder queryString2 = new StringBuilder("delete from contact_detail where 1=1 ");

                    if (contactDetailForm.ContactId != null && !contactDetailForm.ContactId.Equals(""))
                    {
                        queryString2.Append(" and CONTACT_ID=@CONTACT_ID ");
                    }
                    if (contactDetailForm.Name != null && !contactDetailForm.Name.Equals(""))
                    {
                        queryString2.Append(" and NAME=@NAME ");
                    }
                    if (contactDetailForm.Title != null && !contactDetailForm.Title.Equals(""))
                    {
                        queryString2.Append(" and TITLE=@TITLE ");
                    }
                    if (contactDetailForm.CompanyName != null && !contactDetailForm.CompanyName.Equals(""))
                    {
                        queryString2.Append(" and COMPANY_NAME=@COMPANY_NAME ");
                    }
                    if (contactDetailForm.Tel != null && !contactDetailForm.Tel.Equals(""))
                    {
                        queryString2.Append(" and TEL=@TEL ");
                    }
                    if (contactDetailForm.Mobile != null && !contactDetailForm.Mobile.Equals(""))
                    {
                        queryString2.Append(" and MOBILE_NO=@MOBILE_NO ");
                    }
                    if (contactDetailForm.Email != null && !contactDetailForm.Email.Equals(""))
                    {
                        queryString2.Append(" and EMAIL=@EMAIL ");
                    }
                    if (contactDetailForm.Website != null && !contactDetailForm.Website.Equals(""))
                    {
                        queryString2.Append(" and WEBSITE=@WEBSITE ");
                    }
                    if (contactDetailForm.Remark1 != null && !contactDetailForm.Remark1.Equals(""))
                    {
                        queryString2.Append(" and REMARK1=@REMARK1 ");
                    }
                    if (contactDetailForm.Remark2 != null && !contactDetailForm.Remark2.Equals(""))
                    {
                        queryString2.Append(" and REMARK2=@REMARK2 ");
                    }
                    if (contactDetailForm.Remark3 != null && !contactDetailForm.Remark3.Equals(""))
                    {
                        queryString2.Append(" and REMARK3=@REMARK3 ");
                    }
                    if (contactDetailForm.Remark4 != null && !contactDetailForm.Remark4.Equals(""))
                    {
                        queryString2.Append(" and REMARK4=@REMARK4 ");
                    }
                    if (contactDetailForm.Remark5 != null && !contactDetailForm.Remark5.Equals(""))
                    {
                        queryString2.Append(" and REMARK5=@REMARK5 ");
                    }
                    if (contactDetailForm.Remark6 != null && !contactDetailForm.Remark6.Equals(""))
                    {
                        queryString2.Append(" and REMARK6=@REMARK6 ");
                    }
                    if (contactDetailForm.Remark7 != null && !contactDetailForm.Remark7.Equals(""))
                    {
                        queryString2.Append(" and REMARK7=@REMARK7 ");
                    }
                    if (contactDetailForm.Remark8 != null && !contactDetailForm.Remark8.Equals(""))
                    {
                        queryString2.Append(" and REMARK8=@REMARK8 ");
                    }
                    if (contactDetailForm.Remark9 != null && !contactDetailForm.Remark9.Equals(""))
                    {
                        queryString2.Append(" and REMARK9=@REMARK9 ");
                    }
                    if (contactDetailForm.Remark10 != null && !contactDetailForm.Remark10.Equals(""))
                    {
                        queryString2.Append(" and REMARK10=@REMARK10 ");
                    }


                    cmd = new SqlCommand(queryString2.ToString(), con);

                    con.Open();

                    cmd.Parameters.AddWithValue("@NAME", contactDetailForm.Name != null ? contactDetailForm.Name.Trim() : "");
                    cmd.Parameters.AddWithValue("@TITLE", contactDetailForm.Title != null ? contactDetailForm.Title.Trim() : "");
                    cmd.Parameters.AddWithValue("@COMPANY_NAME", contactDetailForm.CompanyName != null ? contactDetailForm.CompanyName.Trim() : "");
                    cmd.Parameters.AddWithValue("@TEL", contactDetailForm.Tel != null ? contactDetailForm.Tel.Trim() : "");
                    cmd.Parameters.AddWithValue("@MOBILE_NO", contactDetailForm.Mobile != null ? contactDetailForm.Mobile.Trim() : "");
                    cmd.Parameters.AddWithValue("@EMAIL", contactDetailForm.Email != null ? contactDetailForm.Email.Trim() : "");
                    cmd.Parameters.AddWithValue("@WEBSITE", contactDetailForm.Website != null ? contactDetailForm.Website.Trim() : "");
                    cmd.Parameters.AddWithValue("@REMARK1", contactDetailForm.Remark1 != null ? contactDetailForm.Remark1.Trim() : "");
                    cmd.Parameters.AddWithValue("@REMARK2", contactDetailForm.Remark2 != null ? contactDetailForm.Remark2.Trim() : "");
                    cmd.Parameters.AddWithValue("@REMARK3", contactDetailForm.Remark3 != null ? contactDetailForm.Remark3.Trim() : "");
                    cmd.Parameters.AddWithValue("@REMARK4", contactDetailForm.Remark4 != null ? contactDetailForm.Remark4.Trim() : "");
                    cmd.Parameters.AddWithValue("@REMARK5", contactDetailForm.Remark5 != null ? contactDetailForm.Remark5.Trim() : "");
                    cmd.Parameters.AddWithValue("@REMARK6", contactDetailForm.Remark6 != null ? contactDetailForm.Remark6.Trim() : "");
                    cmd.Parameters.AddWithValue("@REMARK7", contactDetailForm.Remark7 != null ? contactDetailForm.Remark7.Trim() : "");
                    cmd.Parameters.AddWithValue("@REMARK8", contactDetailForm.Remark8 != null ? contactDetailForm.Remark8.Trim() : "");
                    cmd.Parameters.AddWithValue("@REMARK9", contactDetailForm.Remark9 != null ? contactDetailForm.Remark9.Trim() : "");
                    cmd.Parameters.AddWithValue("@REMARK10", contactDetailForm.Remark10 != null ? contactDetailForm.Remark10.Trim() : "");
                    cmd.Parameters.AddWithValue("@CONTACT_ID", contactDetailForm.ContactId != null ? contactDetailForm.ContactId.Trim() : "");


                    cmd.ExecuteNonQuery();

                    con.Close();
                }

            }
            catch (Exception e)
            {
                return contactDetails;
            }
            finally
            {
                    if (con != null)
                    {
                        con.Close();
                    }
            }
            return contactDetails;
         }


        // GET api/contactdetail
        [HttpPost("generateexcel")]
        public String generateExcel([FromBody] ContactDetail contactDetailForm)
        {
            String filePath="";
            SqlConnection con = new SqlConnection(dataConnection);
            try
            {

                StringBuilder queryString = new StringBuilder("select CONTACT_ID,NAME,TITLE,COMPANY_NAME,TEL,MOBILE_NO,EMAIL,WEBSITE," +
                       " REMARK1,REMARK2,REMARK3,REMARK4,REMARK5,REMARK6,REMARK7,REMARK8,REMARK9,REMARK10,ENT_DATE,MODIFY_DATE " +
                       " from contact_detail where 1=1");

                if (contactDetailForm.ContactId != null && !contactDetailForm.ContactId.Equals(""))
                {
                    queryString.Append(" and CONTACT_ID=@CONTACT_ID ");
                }
                if (contactDetailForm.Name != null && !contactDetailForm.Name.Equals(""))
                {
                    queryString.Append(" and NAME=@NAME ");
                }
                if (contactDetailForm.Title != null && !contactDetailForm.Title.Equals(""))
                {
                    queryString.Append(" and TITLE=@TITLE ");
                }
                if (contactDetailForm.CompanyName != null && !contactDetailForm.CompanyName.Equals(""))
                {
                    queryString.Append(" and COMPANY_NAME=@COMPANY_NAME ");
                }
                if (contactDetailForm.Tel != null && !contactDetailForm.Tel.Equals(""))
                {
                    queryString.Append(" and TEL=@TEL ");
                }
                if (contactDetailForm.Mobile != null && !contactDetailForm.Mobile.Equals(""))
                {
                    queryString.Append(" and MOBILE_NO=@MOBILE_NO ");
                }
                if (contactDetailForm.Email != null && !contactDetailForm.Email.Equals(""))
                {
                    queryString.Append(" and EMAIL=@EMAIL ");
                }
                if (contactDetailForm.Website != null && !contactDetailForm.Website.Equals(""))
                {
                    queryString.Append(" and WEBSITE=@WEBSITE ");
                }
                if (contactDetailForm.Remark1 != null && !contactDetailForm.Remark1.Equals(""))
                {
                    queryString.Append(" and REMARK1=@REMARK1 ");
                }
                if (contactDetailForm.Remark2 != null && !contactDetailForm.Remark2.Equals(""))
                {
                    queryString.Append(" and REMARK2=@REMARK2 ");
                }
                if (contactDetailForm.Remark3 != null && !contactDetailForm.Remark3.Equals(""))
                {
                    queryString.Append(" and REMARK3=@REMARK3 ");
                }
                if (contactDetailForm.Remark4 != null && !contactDetailForm.Remark4.Equals(""))
                {
                    queryString.Append(" and REMARK4=@REMARK4 ");
                }
                if (contactDetailForm.Remark5 != null && !contactDetailForm.Remark5.Equals(""))
                {
                    queryString.Append(" and REMARK5=@REMARK5 ");
                }
                if (contactDetailForm.Remark6 != null && !contactDetailForm.Remark6.Equals(""))
                {
                    queryString.Append(" and REMARK6=@REMARK6 ");
                }
                if (contactDetailForm.Remark7 != null && !contactDetailForm.Remark7.Equals(""))
                {
                    queryString.Append(" and REMARK7=@REMARK7 ");
                }
                if (contactDetailForm.Remark8 != null && !contactDetailForm.Remark8.Equals(""))
                {
                    queryString.Append(" and REMARK8=@REMARK8 ");
                }
                if (contactDetailForm.Remark9 != null && !contactDetailForm.Remark9.Equals(""))
                {
                    queryString.Append(" and REMARK9=@REMARK9 ");
                }
                if (contactDetailForm.Remark10 != null && !contactDetailForm.Remark10.Equals(""))
                {
                    queryString.Append(" and REMARK10=@REMARK10 ");
                }


                SqlCommand cmd = new SqlCommand(queryString.ToString(), con);

                cmd.Parameters.AddWithValue("@CONTACT_ID", contactDetailForm.ContactId != null ? contactDetailForm.ContactId.Trim() : "");
                cmd.Parameters.AddWithValue("@NAME", contactDetailForm.Name != null ? contactDetailForm.Name.Trim() : "");
                cmd.Parameters.AddWithValue("@TITLE", contactDetailForm.Title != null ? contactDetailForm.Title.Trim() : "");
                cmd.Parameters.AddWithValue("@COMPANY_NAME", contactDetailForm.CompanyName != null ? contactDetailForm.CompanyName.Trim() : "");
                cmd.Parameters.AddWithValue("@TEL", contactDetailForm.Tel != null ? contactDetailForm.Tel.Trim() : "");
                cmd.Parameters.AddWithValue("@MOBILE_NO", contactDetailForm.Mobile != null ? contactDetailForm.Mobile.Trim() : "");
                cmd.Parameters.AddWithValue("@EMAIL", contactDetailForm.Email != null ? contactDetailForm.Email.Trim() : "");
                cmd.Parameters.AddWithValue("@WEBSITE", contactDetailForm.Website != null ? contactDetailForm.Website.Trim() : "");
                cmd.Parameters.AddWithValue("@REMARK1", contactDetailForm.Remark1 != null ? contactDetailForm.Remark1.Trim() : "");
                cmd.Parameters.AddWithValue("@REMARK2", contactDetailForm.Remark2 != null ? contactDetailForm.Remark2.Trim() : "");
                cmd.Parameters.AddWithValue("@REMARK3", contactDetailForm.Remark3 != null ? contactDetailForm.Remark3.Trim() : "");
                cmd.Parameters.AddWithValue("@REMARK4", contactDetailForm.Remark4 != null ? contactDetailForm.Remark4.Trim() : "");
                cmd.Parameters.AddWithValue("@REMARK5", contactDetailForm.Remark5 != null ? contactDetailForm.Remark5.Trim() : "");
                cmd.Parameters.AddWithValue("@REMARK6", contactDetailForm.Remark6 != null ? contactDetailForm.Remark6.Trim() : "");
                cmd.Parameters.AddWithValue("@REMARK7", contactDetailForm.Remark7 != null ? contactDetailForm.Remark7.Trim() : "");
                cmd.Parameters.AddWithValue("@REMARK8", contactDetailForm.Remark8 != null ? contactDetailForm.Remark8.Trim() : "");
                cmd.Parameters.AddWithValue("@REMARK9", contactDetailForm.Remark9 != null ? contactDetailForm.Remark9.Trim() : "");
                cmd.Parameters.AddWithValue("@REMARK10", contactDetailForm.Remark10 != null ? contactDetailForm.Remark10.Trim() : "");

                con.Open();

                SqlDataReader rdr = cmd.ExecuteReader();
         
                DataTable dtCustomer = new DataTable("Contact Detail");

                dtCustomer.Columns.Add(new DataColumn("CONTACT_ID"));
                dtCustomer.Columns.Add(new DataColumn("NAME"));
                dtCustomer.Columns.Add(new DataColumn("TITLE"));
                dtCustomer.Columns.Add(new DataColumn("COMPANY_NAME"));
                dtCustomer.Columns.Add(new DataColumn("TEL"));
                dtCustomer.Columns.Add(new DataColumn("MOBILE_NO"));
                dtCustomer.Columns.Add(new DataColumn("EMAIL"));
                dtCustomer.Columns.Add(new DataColumn("WEBSITE"));
                dtCustomer.Columns.Add(new DataColumn("REMARK1"));
                dtCustomer.Columns.Add(new DataColumn("REMARK2"));
                dtCustomer.Columns.Add(new DataColumn("REMARK3"));
                dtCustomer.Columns.Add(new DataColumn("REMARK4"));
                dtCustomer.Columns.Add(new DataColumn("REMARK5"));
                dtCustomer.Columns.Add(new DataColumn("REMARK6"));
                dtCustomer.Columns.Add(new DataColumn("REMARK7"));
                dtCustomer.Columns.Add(new DataColumn("REMARK8"));
                dtCustomer.Columns.Add(new DataColumn("REMARK9"));
                dtCustomer.Columns.Add(new DataColumn("REMARK10"));


                DataRow drItem = dtCustomer.NewRow();

                drItem["CONTACT_ID"] = "CONTACT_ID";
                drItem["NAME"] = "NAME";
                drItem["TITLE"] = "TITLE";
                drItem["COMPANY_NAME"] = "COMPANY_NAME";
                drItem["TEL"] = "TEL";
                drItem["MOBILE_NO"] = "MOBILE_NO";
                drItem["EMAIL"] = "EMAIL";
                drItem["WEBSITE"] = "WEBSITE";
                drItem["REMARK1"] = "REMARK1";
                drItem["REMARK2"] = "REMARK2";
                drItem["REMARK3"] = "REMARK3";
                drItem["REMARK4"] = "REMARK4";
                drItem["REMARK5"] = "REMARK5";
                drItem["REMARK6"] = "REMARK6";
                drItem["REMARK7"] = "REMARK7";
                drItem["REMARK8"] = "REMARK8";
                drItem["REMARK9"] = "REMARK9";
                drItem["REMARK10"] = "REMARK10";

                dtCustomer.Rows.Add(drItem);

                while (rdr.Read())
                {
                   
                    drItem = dtCustomer.NewRow();

                    drItem["CONTACT_ID"] = rdr["CONTACT_ID"]??"-";
                    drItem["NAME"] = rdr["NAME"] ?? "-";
                    drItem["TITLE"] = rdr["TITLE"] ?? "-";
                    drItem["COMPANY_NAME"] = rdr["COMPANY_NAME"] ?? "-";
                    drItem["TEL"] = rdr["TEL"] ?? "-";
                    drItem["MOBILE_NO"] = rdr["MOBILE_NO"] ?? "-";
                    drItem["EMAIL"] = rdr["EMAIL"] ?? "-";
                    drItem["WEBSITE"] = rdr["WEBSITE"] ?? "-";
                    drItem["REMARK1"] = rdr["REMARK1"] ?? "-";
                    drItem["REMARK2"] = rdr["REMARK2"] ?? "-";
                    drItem["REMARK3"] = rdr["REMARK3"] ?? "-";
                    drItem["REMARK4"] = rdr["REMARK4"] ?? "-";
                    drItem["REMARK5"] = rdr["REMARK5"] ?? "-";
                    drItem["REMARK6"] = rdr["REMARK6"] ?? "-";
                    drItem["REMARK7"] = rdr["REMARK7"] ?? "-";
                    drItem["REMARK8"] = rdr["REMARK8"] ?? "-";
                    drItem["REMARK9"] = rdr["REMARK9"] ?? "-";
                    drItem["REMARK10"] = rdr["REMARK10"] ?? "-";

                    dtCustomer.Rows.Add(drItem);
                                      
                }

                con.Close();

                String fileName = "card_excel" + DateTime.Now.ToString("yyyyMMddHHmmssffff") + ".xls";

                filePath = excelPath+ fileName;

                ExcelCreator.Create(dtCustomer, filePath);

                Task<String> t2 = uploadFile(filePath, fileName);

                filePath = t2.Result.ToString();
             
            }
            catch (Exception e)
            {
                return filePath;
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                }
            }
            return filePath;
        }

        private static async Task<String>  uploadFile(String filePath,String fileName)
        {
            String azurefilePath="";

            CloudStorageAccount cloudStorageAccount = CloudStorageAccount.Parse(storageConnection);

            //create a block blob 

            CloudBlobClient cloudBlobClient = cloudStorageAccount.CreateCloudBlobClient();

            //create a container 

            CloudBlobContainer cloudBlobContainer = cloudBlobClient.GetContainerReference("excelfilecontainer");

            //create a container if it is not already exists

            if (await cloudBlobContainer.CreateIfNotExistsAsync())
            {

                await cloudBlobContainer.SetPermissionsAsync(new BlobContainerPermissions
                { PublicAccess = BlobContainerPublicAccessType.Blob });

            }

            //get Blob reference

            CloudBlockBlob cloudBlockBlob = cloudBlobContainer.GetBlockBlobReference(fileName);
                      

            using (var fileStream = System.IO.File.OpenRead(filePath))
            {
              await cloudBlockBlob.UploadFromStreamAsync(fileStream);
            }

            azurefilePath = azureServerPath + fileName;

            return azurefilePath;
        }
    }
}
