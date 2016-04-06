using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using SpDemoProviderWeb.Models;

namespace SpDemoProviderWeb.Controllers
{
    public class HomeController : Controller
    {
        [SharePointContextFilter]
        public ActionResult Index()
        {

            var spContext = SharePointContextProvider.Current.GetSharePointContext(HttpContext);

            var data = new List<TelefonoViewModel>();
            using (var clientContext = spContext.CreateUserClientContextForSPHost())
            {
                if (clientContext != null)
                {
                    //spUser = clientContext.Web.CurrentUser;
                    //clientContext.Load(spUser, user => user.Title);
                    //clientContext.ExecuteQuery();
                    //ViewBag.UserName = spUser.Title;

                    var telefonoList = clientContext.Web.Lists.GetByTitle("Telefonos");
                    clientContext.Load(telefonoList);
                    clientContext.ExecuteQuery();

                    var query = new CamlQuery();
                    var telefonosItems = telefonoList.GetItems(query);
                    clientContext.Load(telefonosItems);
                    clientContext.ExecuteQuery();

                    foreach (var telefonosItem in telefonosItems)
                    {
                        data.Add(TelefonoViewModel.FromListItem(telefonosItem));
                    }
                }
            }

            return View(data);
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }

        public ActionResult Add()
        {
            return View(new TelefonoViewModel());
        }

        [HttpPost]
        public ActionResult Add(TelefonoViewModel model)
        {
            var spContext = SharePointContextProvider.Current.GetSharePointContext(HttpContext);

            using (var clientContext = spContext.CreateUserClientContextForSPHost())
            {
                if (clientContext != null)
                {
                    var telefonoList = clientContext.Web.Lists.GetByTitle("Telefonos");
                    clientContext.Load(telefonoList);
                    clientContext.ExecuteQuery();

                    var listCreationInfo = new ListItemCreationInformation();
                    var item = telefonoList.AddItem(listCreationInfo);
                    item["Title"] = model.Nombre;
                    item["Numero"] = model.Numero;
                    item.Update();
                    clientContext.ExecuteQuery();


                }
            }

            return RedirectToAction("Index", new { SPHostUrl = SharePointContext.GetSPHostUrl(HttpContext.Request).AbsoluteUri });
        }

        public ActionResult Delete(int id)
        {
            var spContext = SharePointContextProvider.Current.GetSharePointContext(HttpContext);

            var data = new List<TelefonoViewModel>();
            using (var clientContext = spContext.CreateUserClientContextForSPHost())
            {
                if (clientContext != null)
                {

                    var telefonoList = clientContext.Web.Lists.GetByTitle("Telefonos");
                    var telefonosItem = telefonoList.GetItemById(id);

                    telefonosItem.DeleteObject();
                    clientContext.ExecuteQuery();

                }
            }

            return RedirectToAction("Index", new { SPHostUrl = SharePointContext.GetSPHostUrl(HttpContext.Request).AbsoluteUri });
        }

        public ActionResult Update(int id)
        {
            var spContext = SharePointContextProvider.Current.GetSharePointContext(HttpContext);
            TelefonoViewModel model = null;
            var data = new List<TelefonoViewModel>();
            using (var clientContext = spContext.CreateUserClientContextForSPHost())
            {
                if (clientContext != null)
                {

                    var telefonoList = clientContext.Web.Lists.GetByTitle("Telefonos");
                    clientContext.Load(telefonoList);
                    var telefonosItem = telefonoList.GetItemById(id);
                    clientContext.Load(telefonosItem);
                    clientContext.ExecuteQuery();
                    model = TelefonoViewModel.FromListItem(telefonosItem);
                }
            }

            return View(model);
        }

        [HttpPost]
        public ActionResult Update(TelefonoViewModel model)
        {
            var spContext = SharePointContextProvider.Current.GetSharePointContext(HttpContext);
            var data = new List<TelefonoViewModel>();
            using (var clientContext = spContext.CreateUserClientContextForSPHost())
            {
                if (clientContext != null)
                {

                    var telefonoList = clientContext.Web.Lists.GetByTitle("Telefonos");
                    var telefonosItem = telefonoList.GetItemById(model.Id);

                    telefonosItem["Title"] = model.Nombre;
                    telefonosItem["Numero"] = model.Numero;
                    telefonosItem.Update();
                    clientContext.ExecuteQuery();
                }
            }
            return RedirectToAction("Index", new { SPHostUrl = SharePointContext.GetSPHostUrl(HttpContext.Request).AbsoluteUri });

        }

    }
}
