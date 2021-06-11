
using Ibrahim.OutlookApi;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Ibrahim.Test
{
    class Program
    {
        static void Main(string[] args)
        {
            OutlookHelper helper = new OutlookHelper();

            //Aşağdaki kodu test etmeden önce Outlook içerisinde "Kişiler" klasörüne 
            //"Yazılım Eğitimleri" isminde bir klasör oluşturulmalıdır.
            if (helper.OzelKlasorVarMi("Yazılım Eğitimleri"))
            {
                //Aşağıdaki kod "Yazılım Eğitimleri" klasörü altında "Yazılım Uzmanlığı Eğitimi"
                //isminde bir klasör oluşturulmaktadır.
                bool ozelKlasorBasarili=helper.OzelKlasorOlustur("Yazılım Uzmanlığı Eğitimi", "Yazılım Eğitimleri");
                //Özel klasörümüz başarıyla oluşturulduysa
                if (ozelKlasorBasarili)
                {
                    bool dagitimListesiTamam= helper.DagitimListesiOlustur("11 Haziran Grubu", "Yazılım Uzmanlığı Eğitimi");
                    //"11 Haziran Grubu" isimli dağıtım listesi başarıyla oluşturuldu ise
                    if (dagitimListesiTamam)
                    {
                        //Aşağıdaki kod yardımıyla "Yazılım Uzmanlığı Eğitimi" isimli
                        //klasörde yer alan "11 Haziran Grubu" isimli dağıtım listesine
                        //yeni bir kişi eklemektedir.
                        helper.ListeyeKisiEkle("Yazılım Uzmanlığı Eğitimi", "11 Haziran Grubu", new Kisi
                        {
                            EpostaAdresi="test@gmail.com",
                            GorunenAd="Ahmet Mehmet"
                        });
                    }
                }
            }

            Console.WriteLine("İşlem tamam");
            Console.ReadKey();
        }
    }
}
