using Microsoft.AspNetCore.Mvc;
using AcuMenough.Models;
using FluentValidation.Results;
using System.Collections.Generic;
using OfficeOpenXml;

namespace AcuMenough.Controllers
{
    public class LocationController : Controller
    {
        private const string ExcelFilePath = @"floor.xlsx";


        public ActionResult Index()
        {
            List<Location> locations = ReadLocationsFromExcel();
            return View(locations);
        }

        public ActionResult Create()
        {
            return View();
        }

        [HttpPost]
        public ActionResult Create(Location location)
        {
            WriteLocationToExcel(location);
            return RedirectToAction("Index");
        }

        public ActionResult Edit(string locationName)
        {
            Location location = ReadLocationsFromExcel().FirstOrDefault(l => l.LocationName == locationName);
            return View(location);
        }

        [HttpPost]
        public ActionResult Edit(Location location)
        {
            UpdateLocationInExcel(location);
            return RedirectToAction("Index");
        }

        public ActionResult Delete(string locationName)
        {
            Location location = ReadLocationsFromExcel().FirstOrDefault(l => l.LocationName == locationName);
            return View(location);
        }


        [HttpPost]
        public ActionResult DeleteConfirmed(string locationName)
        {
            DeleteLocationFromExcel(locationName);
            return RedirectToAction("Index");
        }

        private List<Location> ReadLocationsFromExcel()
        {
            List<Location> locations = new List<Location>();
            //var fileIfo= new FileInfo(ExcelFilePath);
            using (ExcelPackage package = new ExcelPackage(new FileInfo(ExcelFilePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                int rowCount = worksheet.Dimension.Rows;

                for (int row = 2; row <= rowCount; row++)
                {
                    locations.Add(new Location
                    {
                        LocationName = worksheet.Cells[row, 1].Value.ToString(),
                        LocationId = int.Parse(worksheet.Cells[row, 2].Value.ToString()),
                        IsClearance = worksheet.Cells[row, 3].Value.ToString()
                    });
                }
            }

            return locations;
        }

        private void WriteLocationToExcel(Location location)
        {
            using (ExcelPackage package = new ExcelPackage(new FileInfo(ExcelFilePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                int rowCount = worksheet.Dimension.Rows;

                worksheet.Cells[rowCount + 1, 1].Value = location.LocationName;
                worksheet.Cells[rowCount + 1, 2].Value = location.LocationId;
                worksheet.Cells[rowCount + 1, 3].Value = location.IsClearance;

                package.Save();
            }
        }

        private void UpdateLocationInExcel(Location location)
        {
            using (ExcelPackage package = new ExcelPackage(new FileInfo(ExcelFilePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                int rowCount = worksheet.Dimension.Rows;
                for (int row = 2; row <= rowCount; row++)
                {
                    if (worksheet.Cells[row, 1].Value.ToString() == location.LocationName)
                    {
                        worksheet.Cells[row, 2].Value = location.LocationId;
                        worksheet.Cells[row, 3].Value = location.IsClearance;
                        break;
                    }
                }

                package.Save();
            }
        }

        private void DeleteLocationFromExcel(string locationName)
        {
            using (ExcelPackage package = new ExcelPackage(new FileInfo(ExcelFilePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                int rowCount = worksheet.Dimension.Rows;
                for (int row = 2; row <= rowCount; row++)
                {
                    if (worksheet.Cells[row, 1].Value.ToString() == locationName)
                    {
                        worksheet.DeleteRow(row);
                        break;
                    }
                }

                package.Save();
            }
        }
    }
}