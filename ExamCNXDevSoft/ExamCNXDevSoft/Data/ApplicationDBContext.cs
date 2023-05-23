using ExamCNXDevSoft.Models;
using Microsoft.EntityFrameworkCore;

namespace ExamCNXDevSoft.Data
{
    public class ApplicationDBContext:DbContext
    {
        public ApplicationDBContext(DbContextOptions<ApplicationDBContext> options):base(options) { 

        }
        public DbSet<Student> Students { get; set; }
    }
}
