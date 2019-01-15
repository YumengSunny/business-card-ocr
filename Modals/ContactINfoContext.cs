using Microsoft.EntityFrameworkCore;
using BCardReader.Modals;

namespace BCardReader.Models
{
    public class ContactInfoContex : DbContext
    {
        public ContactInfoContex(DbContextOptions<ContactInfoContex> options)
            : base(options)
        {
        }

        public DbSet<ContactInfo> ContactInfoItems { get; set; }
    }
}