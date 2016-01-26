using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using MVCDisplayWebGrid.Models;
using System.Web.Helpers;

namespace MVCDisplayWebGrid.Controllers
{
    public class EmployeeController : Controller
    {
        public ActionResult Index()
        {
            EmployeeDBContext db = new EmployeeDBContext();
            return View(db.Employees.ToList());
        }

        public void ExportData()
        {
            EmployeeDBContext db = new EmployeeDBContext();
            List<Employee> employees = new List<Employee>();
            employees = db.Employees.ToList();

            WebGrid webGrid = new WebGrid(employees, canPage: true, rowsPerPage: 10);
            string gridData = webGrid.GetHtml(
                columns: webGrid.Columns(
                webGrid.Column(columnName: "Name", header: "Name"),
                webGrid.Column(columnName: "Designation", header: "Designation"),
                webGrid.Column(columnName: "Gender", header: "Gender"),
                webGrid.Column(columnName: "Salary", header: "Salary"),
                webGrid.Column(columnName: "City", header: "City"),
                webGrid.Column(columnName: "State", header: "State"),
                webGrid.Column(columnName: "Zip", header: "Zip")
            )).ToString();

            Response.ClearContent();
            Response.AddHeader("content-disposition","attachment; filename=Employee-Details.xls");
            Response.ContentType = "applicatiom/excel";
            Response.Write(gridData);
            Response.End();
        }
	}
}