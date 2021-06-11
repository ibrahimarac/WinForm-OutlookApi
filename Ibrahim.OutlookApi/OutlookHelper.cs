
using NetOffice;
using Outlook = NetOffice.OutlookApi;
using NetOffice.OutlookApi.Enums;
using NetOffice.OutlookApi;
using System.Collections.Generic;
using System;
using System.Text;

namespace Ibrahim.OutlookApi
{
    public class OutlookHelper
    {
        #region Kullanılan Private Referanslar

        Outlook.Application outlookApp;
        List<MAPIFolder> folders;

        #endregion

        #region Kurucu Metod

        public OutlookHelper()
        {
            outlookApp = new Outlook.Application();
            folders = new List<MAPIFolder>();
        }

        #endregion

        #region Aşağıda kullandığımız private metodlar

        MAPIFolder KlasorBul(_Folders bakilacakKlasorler, string klasorAdi)
        {
            //Aşağıdaki recursive metod yardımıyla gelen klasörler ve
            //varsa onların alt klasörlerinde <klasorAdi> isimli klasör
            //var mı diye bakıyoruz.
            foreach (var item in bakilacakKlasorler)
            {
                //Klasörü bulduysak metoddan çıkıyoruz.
                if (item.Name == klasorAdi)
                    return item;
            }
            //aranan klasör bulunamadıysa bu klasörlerden alt klasöre sahip
            //olanların alt klasörlerinde de arıyoruz.
            foreach (var item in bakilacakKlasorler)
            {
                if (item.Folders.Count != 0)
                    return KlasorBul(item.Folders, klasorAdi);
            }
            return null;
        }

        void AltKlasorleriDoldur(MAPIFolder folder)
        {
            //Gelen klasörün tüm alt klasörlerini dolduran metod
            foreach (var item in folder.Folders)
            {
                folders.Add(item);
                if (item.Folders.Count != 0)
                    AltKlasorleriDoldur(item);
            }
        }

        bool DagitimListesiVarmi(string listeAdi, string klasorAdi)
        {
            //Gönderilen <listeAdi> isimli dağıtım listesinin <klasorAdi>
            //isimli klasörde olup olmadığına bakıyoruz.
            bool dagitimListesiVar = false;
            MAPIFolder fldContacts = (MAPIFolder)outlookApp.Session.GetDefaultFolder(OlDefaultFolders.olFolderContacts);
            MAPIFolder folder = KlasorBul(fldContacts.Folders, klasorAdi);
            foreach (var item in folder.Items)
            {
                if (item is DistListItem && (item as DistListItem).DLName == listeAdi)
                {
                    dagitimListesiVar = true;
                    break;
                }
            }
            return dagitimListesiVar;
        }

        DistListItem DagitimListesiGetir(string klasorAdi, string listeAdi)
        {
            //Bu metod gelen <klasorAdi> isimli klasörde <listeAdi> isimli bir listeyi (dağıtım listesi)
            //arıyor. Bulduğu takdirde bu listeyi DistListItem türünden gönderiyor.
            DistListItem liste = null;
            //Klasörde böyle bir dağıtım listesi varsa listeyi elde ederek gönderelim.
            if (DagitimListesiVarmi(listeAdi, klasorAdi))
            {
                MAPIFolder fldContacts = (MAPIFolder)outlookApp.Session.GetDefaultFolder(OlDefaultFolders.olFolderContacts);
                MAPIFolder folder = KlasorBul(fldContacts.Folders, klasorAdi);
                foreach (var item in folder.Items)
                {
                    if (item is DistListItem && (item as DistListItem).DLName == listeAdi)
                    {
                        liste = item as DistListItem;
                        break;
                    }
                }
            }
            return liste;
        }

        #endregion

        #region Klasör İşlemleri

        public bool OzelKlasorVarMi(string klasor)
        {
            //<klasor> isimli klasörün Outlook'da olup olmadığına bakıyor.
            MAPIFolder fldContacts = (MAPIFolder)outlookApp.Session.GetDefaultFolder(OlDefaultFolders.olFolderContacts);
            MAPIFolder folder = KlasorBul(fldContacts.Folders, klasor);
            return folder != null;
        }
                
        public bool OzelKlasorOlustur(string klasorAdi, string ustKlasorAdi)
        {
            //<ustKlasorAdi> isimli klasörün altında <klasorAdi> isimli klasörü oluşturmaya
            //çalışıyoruz.
            MAPIFolder fldContacts = (MAPIFolder)outlookApp.Session.GetDefaultFolder(OlDefaultFolders.olFolderContacts);
            //Gönderilen <ustKlasorAdi> isimli bir klasör var mı? Eğer böyle bir üst klasör yoksa
            //<klasorAdi> isimli klasörü bu klasör içerisine açamayız.
            MAPIFolder folder = KlasorBul(fldContacts.Folders, ustKlasorAdi);
            if (folder != null)//üst klasör bulundu
                folder.Folders.Add(klasorAdi);
            return folder != null;
        }

        public bool OzelKlasorSil(string klasorAdi)
        {
            MAPIFolder fldContacts = (MAPIFolder)outlookApp.Session.GetDefaultFolder(OlDefaultFolders.olFolderContacts);
            //Tüm klasörlerde <klasorAdi> isimli klasörü arıyorum. Eğer bulursam silme işlemine geçiyorum.
            MAPIFolder folder = KlasorBul(fldContacts.Folders, klasorAdi);
            if (folder != null)
                folder.Delete();
            return folder != null;
        }

        public bool OzelKlasorAdiGuncelle(string klasorAdi, string yeniKlasorAdi, string ustKlasorAdi)
        {
            MAPIFolder fldContacts = (MAPIFolder)outlookApp.Session.GetDefaultFolder(OlDefaultFolders.olFolderContacts);
            //Eski klasör adıyla bir klasör var mı
            MAPIFolder folder = KlasorBul(fldContacts.Folders, klasorAdi);
            if (folder != null)
            {
                //eski klasör adıyla bir klasör bulundu. Klasör adını değiştir.
                folder.Name = yeniKlasorAdi;
            }
            return folder != null;
        }

        public List<MAPIFolder> OzelKlasorleriGetir()
        {
            //folders isimli List<MAPIFolder> türünden koleksiyonumuzu temizliyoruz.
            folders.Clear();
            //Default klasöründeki üyelere erişiyoruz.
            MAPIFolder fldContacts = (MAPIFolder)outlookApp.Session.GetDefaultFolder(OlDefaultFolders.olFolderContacts);
            //Default klasörün alt klasörlerini dolduruyoruz. Bu metod tüm klasörleri getirir.
            AltKlasorleriDoldur(fldContacts);
            return folders;
        }

        #endregion

        #region Dağıtım Listesi İşlemleri

        public bool DagitimListesiOlustur(string listeAdi, string klasorAdi)
        {
            //Bu metod <listeAdi> isimli dağıtım listesini, <klasorAdi> isimli
            //klasöre eklemektedir.

            //Klasörde böyle bir dağıtım listesi var mı?
            //Dağıtım listesi yoksa oluşturalım.
            if (!DagitimListesiVarmi(listeAdi, klasorAdi))
            {
                //Dağıtım listesi yoksa oluşturalım.

                //Default klasörüne ulaşıyoruz.
                MAPIFolder fldContacts = (MAPIFolder)outlookApp.Session.GetDefaultFolder(OlDefaultFolders.olFolderContacts);
                //Default klasörü altında <klasorAdi> isimli klasörü arıyoruz.
                MAPIFolder folder = KlasorBul(fldContacts.Folders, klasorAdi);
                if (folder != null)
                {
                    //<klasorAdi> isimli klasöre dağıtım listesini oluşturuyoruz.
                    DistListItem distList = folder.Items.Add(OlItemType.olDistributionListItem)
                                                                      as Outlook.DistListItem;
                    //Dağıtım listesine ismini veriyoruz.
                    distList.DLName = listeAdi;
                    //listeyi kaydediyoruz.
                    distList.Save();
                    return true;
                }
            }
            return false;
        }

        public bool DagitimListesiSil(string listeAdi, string klasorAdi)
        {
            //Bu metod <listeAdi> isimli dağıtım listesini, <klasorAdi> isimli
            //klasörden silmektedir.

            //Klasörde böyle bir dağıtım listesi var mı?
            if (DagitimListesiVarmi(listeAdi, klasorAdi))
            {
                //Dağıtım listesi varsa silelim
                DagitimListesiGetir(klasorAdi, listeAdi).Delete();
                return true;
            }
            return false;
        }

        public bool ListeyeKisiEkle(string klasorAdi, string listeAdi, Kisi kisi)
        {
            //Bu metod <klasorAdi> isimli klasörde yer alan <listeAdi> isimli dağıtım listesine yeni
            //bir kişi eklemeye çalışmaktadır.

            //Klasörde böyle bir dağıtım listesi varsa kişiyi ekle
            if (DagitimListesiVarmi(listeAdi, klasorAdi))
            {
                MAPIFolder fldContacts = (MAPIFolder)outlookApp.Session.GetDefaultFolder(OlDefaultFolders.olFolderContacts);
                MAPIFolder folder = KlasorBul(fldContacts.Folders, klasorAdi);
                //Kişiyi ekleyeceğimiz dağıtım listesine ulaşalım.
                DistListItem distList = DagitimListesiGetir(klasorAdi, listeAdi);
                //Kişiyi oluşturalım.
                //  "Yeni Kişi <test@gmail.com>" gibi bir eklemede Yeni Kişi başlık <> işaretleri arasındaki veri de
                //  eposta olarak kabul edilmektedir.
                Recipient recip = outlookApp.Session.CreateRecipient(string.Format("{0} <{1}>", kisi.GorunenAd, kisi.EpostaAdresi));
                recip.Resolve();
                //listeye ekle
                distList.AddMember(recip);
                return true;
            }
            return false;
        }

        public List<Kisi> ListeKisileriGetir(string klasorAdi, string listeAdi)
        {
            //Bu metod <klasorAdi> isimli klasördeki <listeAdi> isimli listede yer alan kişileri
            //getiren metoddur.

            //kişileri toplayacağımız koleksiyon
            List<Kisi> kisiler = new List<Kisi>();

            //Klasörde böyle bir dağıtım listesi varsa kişileri alalım
            if (DagitimListesiVarmi(listeAdi, klasorAdi))
            {
                MAPIFolder fldContacts = (MAPIFolder)outlookApp.Session.GetDefaultFolder(OlDefaultFolders.olFolderContacts);
                MAPIFolder folder = KlasorBul(fldContacts.Folders, klasorAdi);
                //Dağıtım listesine ulaşalım.
                DistListItem distList = DagitimListesiGetir(klasorAdi, listeAdi);
                //Dağıtım listesinin üyelerine ait index numaraları 1'den başlıyor.
                int index = 1;
                //kişilerin aınması tamamlandığında distList.GetMember(index) null olacaktır.
                while (distList.GetMember(index) != null)
                {
                    //sıradaki kişiyi listeye ekle
                    kisiler.Add(new Kisi
                    {
                        EpostaAdresi = distList.GetMember(index).Address,
                        GorunenAd = distList.GetMember(index).Name
                    });
                    index++;
                }
            }

            return kisiler;
        }

        public bool ListeyeKisileriEkle(string klasorAdi, string listeAdi, List<Kisi> kisiler)
        {
            //<klasorAdi isimli klasörde yer alan <listeAdi> isimli listeye
            //birden fazla kişiyi tek seferde eklemek için kullanılan metoddur.

            //Klasörde böyle bir dağıtım listesi varsa kişiyi ekle
            if (DagitimListesiVarmi(listeAdi, klasorAdi))
            {
                MAPIFolder fldContacts = (MAPIFolder)outlookApp.Session.GetDefaultFolder(OlDefaultFolders.olFolderContacts);
                MAPIFolder folder = KlasorBul(fldContacts.Folders, klasorAdi);
                DistListItem distList = DagitimListesiGetir(klasorAdi, listeAdi);
                //Tüm kişileri sırayla gezerek listeye ekliyoruz.
                foreach (var kisi in kisiler)
                {
                    Recipient recip = outlookApp.Session.CreateRecipient(string.Format("{0} <{1}>", kisi.GorunenAd, kisi.EpostaAdresi));
                    recip.Resolve();
                    distList.AddMember(recip);
                }
                return true;
            }
            return false;
        }

        #endregion

    }
}
