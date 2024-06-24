import os
import pandas as pd
from io import StringIO
from selenium import webdriver
from bs4 import BeautifulSoup

def scrape_to_excel(stock_symbol):
    url = f'https://fintables.com/sirketler/{stock_symbol}/finansal-tablolar/bilanco'
    driver = webdriver.Chrome()
    driver.get(url)

    # JavaScript'in yüklenmesini bekle
    driver.implicitly_wait(20)  # gerekirse süreyi ayarlayın
    html_icerik = driver.page_source
    soup = BeautifulSoup(html_icerik, 'html.parser')
    tables = soup.find_all("table", attrs={"class": "w-full"})
    # İlk tablonun ilk satırından sütun başlıklarını al
    first_table = tables[0]
    headers = first_table.find_all('th')
    column_mapping = {i: header.get_text() for i, header in enumerate(headers)}
    # Tüm tabloları birleştirmek için bir listeye ekleyin
    all_dfs = []
    for table in tables:
        for span in table.find_all("span",class_="absolute left-full font-bold text-[10px] pl-1 "
                                                 "text-shared-danger-solid-01 opacity-0 group-hover:opacity-100"):
            span.decompose()
        for span in table.find_all("span", class_="absolute left-full font-bold text-[10px] pl-1 "
                                                 "text-shared-success-solid-01 opacity-0 group-hover:opacity-100"):
            span.decompose()
        for span in table.find_all("span",class_="absolute left-full font-bold text-[10px] pl-1 text-foreground-01 "
                                                 "opacity-0 group-hover:opacity-100"):
            span.decompose()

        html_string = str(table)
        html_io = StringIO(html_string)
        df = pd.read_html(html_io)[0]

        # Sütunları yeniden adlandır
        if not df.empty:
            df.columns = [column_mapping.get(i, f"Unknown_{i}") for i in range(df.shape[1])]
        all_dfs.append(df)

    # Tüm tabloları alt alta ekle
    combined_df = pd.concat(all_dfs, axis=0, ignore_index=True)

    # Özel bir konum belirleyin.
    save_path = 'Borsa'
    excel_filename = f"{stock_symbol}_data.xlsx"
    full_path = os.path.join(save_path, excel_filename)

    # Birleştirilen tabloyu Excel dosyasına kaydedin
    combined_df.to_excel(full_path, index=False)
    print(f"{full_path} başarıyla kaydedildi.")
"""
HisseListe=[['BINHO'], ['AVOD'], ['A1CAP'], ['ACSEL'], ['ADEL'], ['ADESE'], ['ADGYO'], ['AFYON'], ['AGHOL'], ['AGESA'], ['AGROT'], ['AHGAZ'], ['AKSFA'], ['AKFK'], ['AKMEN'], ['AKBNK'], ['AKCKM'], ['AKCNS'], ['AKDFA'], ['AKYHO'], ['AKENR'], ['AKFGY'], ['AKFYE'], ['ATEKS'], ['AKSGY'], ['AKMGY'], ['AKSA'], ['AKSEN'], ['AKGRT'], ['AKSUE'], ['AKTVK'], ['AKTIF'], ['ALCAR'], ['ALGYO'], ['ALARK'], ['ALBRK'], ['ALCTL'], ['ALFAS'], ['ALJF'], ['ALKIM'], ['ALKA'], ['ALNUS'], ['ALFIN'], ['AYCES'], ['ALTNY'], ['ALKLC'], ['ALMAD'], ['ALVES'], ['ANSGR'], ['AEFES'], ['ANHYT'], ['ASUZU'], ['ANGEN'], ['ANELE'], ['ARCLK'], ['ARDYZ'], ['ARENA'], ['ARNFK'], ['ARSAN'], ['ARTMS'], ['ARZUM'], ['ASGYO'], ['ASELS'], ['ASTOR'], ['ATAGY'], ['ATAYM'], ['ATAKP'], ['AGYO'], ['ATLFA'], ['ATSYH'], ['ATLAS'], ['ATATP'], ['AVGYO'], ['AVTUR'], ['AVHOL'], ['AVPGY'], ['AYDEM'], ['AYEN'], ['AYES'], ['AYGAZ'], ['AZTEK'],
            ['BAGFS'], ['BAKAB'], ['BALAT'], ['BNTAS'], ['BANVT'], ['BARMA'], ['BASGZ'], ['BASCM'], ['BEGYO'], ['BTCIM'], ['BSOKE'], ['BYDNR'], ['BAYRK'], ['BERA'], ['BRKT'], ['BRKSN'], ['BJKAS'], ['BEYAZ'], ['BIENF'], ['BIENY'], ['BLCYT'], ['BLKOM'], ['BIMAS'], ['BIOEN'], ['BRKVY'], ['BRKO'], ['BRLSM'], ['BRMEN'], ['BIZIM'], ['BMSTL'], ['BMSCH'], ['BNPPI'], ['BNPFK'], ['BOBET'], ['BORSK'], ['BORLS'], ['BRSAN'], ['BRYAT'], ['BFREN'], ['BOSSA'], ['BRISA'], ['BURCE'], ['BURVA'], ['BUCIM'], ['BVSAN'], ['BIGCH'], ['CRFSA'], ['CASA'], ['CEOEM'], ['CCOLA'], ['CONSE'], ['COSMO'], ['CRDFA'], ['CVKMD'], ['CWENE'], ['CAGFA'], ['CLDNM'], ['CANTE'], ['CATES'], ['CLEBI'], ['CELHA'], ['CLKMT'], ['CEMAS'], ['CEMTS'], ['CMBTN'], ['CMENT'], ['CIMSA'], ['CUSAN'],
            ['DYBNK'], ['DAGI'], ['DAGHL'], ['DAPGM'], ['DARDL'], ['DGATE'], ['DGRVK'], ['DMSAS'], ['DENGE'], ['DENFA'], ['DNFIN'], ['DZGYO'], ['DZYMK'], ['DENIZ'], ['DERIM'], ['DERHL'], ['DESA'], ['DESPC'], ['DEVA'], ['DNISI'], ['DIRIT'], ['DITAS'], ['DMRGD'], ['DOCO'], ['DOFER'], ['DOBUR'], ['DOHOL'], ['DTRND'], ['DGNMO'], ['ARASE'], ['DOGUB'], ['DGGYO'], ['DOAS'], ['DFKTR'], ['DOKTA'], ['DURDO'], ['DNYVA'], ['DYOBY'], ['EDATA'], ['EBEBK'], ['ECZYT'], ['EDIP'], ['EGEEN'], ['EGGUB'], ['EGPRO'], ['EGSER'], ['EPLAS'], ['ECILC'], ['EKER'], ['EKIZ'], ['EKOFA'], ['EKOS'], ['EKOVR'], ['EKSUN'], ['ELITE'], ['EMKEL'], ['EMNIS'], ['EMIRV'], ['EKTVK'], ['EKGYO'], ['EMVAR'], ['ENJSA'], ['ENERY'], ['ENKAI'], ['ENSRI'], ['ERBOS'], ['ERCB'], ['EREGL'], ['ERGLI'], ['KIMMR'],
            ['ERSU'], ['ESCAR'], ['ESCOM'], ['ESEN'], ['ETILR'], ['EUKYO'], ['EUYO'], ['ETYAT'], ['EUHOL'], ['TEZOL'], ['EUREN'], ['EUPWR'], ['EYGYO'], ['FADE'], ['FSDAT'], ['FMIZP'], ['FENER'], ['FIBAF'], ['FBBNK'], ['FLAP'], ['FONET'], ['FROTO'], ['FORMT'], ['FORTE'], ['FRIGO'], ['FZLGY'], ['GWIND'], ['GSRAY'], ['GAPIN'], ['GARFA'], ['GARFL'], ['GRNYO'], ['GEDIK'], ['GEDZA'], ['GLCVY'], ['GENIL'], ['GENTS'], ['GEREL'], ['GZNMI'], ['GIPTA'], ['GMTAS'], ['GESAN'], ['GLBMD'], ['GLYHO'], ['GGBVK'], ['GSIPD'], ['GOODY'], ['GOKNR'], ['GOLTS'], ['GOZDE'], ['GRTRK'], ['GSDDE'], ['GSDHO'], ['GUBRF'], ['GLRYH'], ['GRSEL'], ['SAHOL'], ['HALKF'], ['HLGYO'], ['HLVKS'], ['HALKI'], ['HRKET'], ['HATSN'], ['HATEK'], ['HDFFL'], ['HDFGS'], ['HEDEF'], ['HEKTS'], ['HKTM'], ['HTTBT'], ['HOROZ'], ['HUBVC'], ['HUNER'], ['HUZFA'], ['HURGZ'], ['ENTRA'],
            ['ICBCT'], ['ICUGS'], ['INGRM'], ['INVEO'], ['INVAZ'], ['INVES'], ['ISKPL'], ['IEYHO'], ['IDGYO'], ['IHEVA'], ['IHLGM'], ['IHGZT'], ['IHAAS'], ['IHLAS'], ['IHYAY'], ['IMASM'], ['INDES'], ['INFO'], ['INTEM'], ['IPEKE'], ['ISDMR'], ['ISTFK'], ['ISFAK'], ['ISFIN'], ['ISGYO'], ['ISGSY'], ['ISMEN'], ['ISYAT'], ['ISBIR'], ['ISSEN'], ['IZINV'], ['IZENR'], ['IZMDC'], ['IZFAS'], ['JANTS'], ['KFEIN'], ['KLKIM'], ['KLSER'], ['KLVKS'], ['KAPLM'], ['KRDMA'], ['KRDMB'], ['KRDMD'], ['KAREL'], ['KARSN'], ['KRTEK'], ['KARYE'], ['KARTN'], ['KATVK'], ['KTLEV'], ['KATMR'], ['KAYSE'], ['KNTFA'], ['KENT'], ['KERVT'], ['KRVGD'], ['KERVN'], ['KZBGY'], ['KLGYO'], ['KLRHO'], ['KMPUR'], ['KLMSN'], ['KCAER'], ['KFKTF'], ['KOCFN'], ['KCHOL'], ['KOCMT'], ['KCSIS'], ['KLSYN'], ['KNFRT'], ['KONTR'], ['KONYA'], ['KONKA'], ['KGYO'], ['KORDS'], ['KRPLS'], ['KORTS'], ['KOTON'], ['KOZAL'], ['KOZAA'], ['KOPOL'], ['KRGYO'], ['KRSTL'], ['KRONT'], ['KTKVK'], ['KSTUR'], ['KUVVA'], ['KUYAS'], ['KBORU'], ['KZGYO'], ['KUTPO'], ['KTSKR'],
            ['LIDER'], ['LIDFA'], ['LILAK'], ['LMKDC'], ['LINK'], ['LOGO'], ['LKMNH'], ['LRSHO'], ['LUKSK'], ['MACKO'], ['MAKIM'], ['MAKTK'], ['MANAS'], ['MRBAS'], ['MAGEN'], ['MRMAG'], ['MARKA'], ['MAALT'], ['MRSHL'], ['MRGYO'], ['MARTI'], ['MTRKS'], ['MAVI'], ['MZHLD'], ['MEDTR'], ['MEGMT'], ['MEGAP'], ['MEKAG'], ['MEKMD'], ['MNDRS'], ['MEPET'], ['MERCN'], ['MRBKF'], ['MBFTR'], ['MERIT'], ['MERKO'], ['METUR'], ['METRO'], ['MTRYO'], ['MHRGY'], ['MIATK'], ['MGROS'], ['MIPAZ'], ['MSGYO'], ['MPARK'], ['MMCAS'], ['MOBTL'], ['MOGAN'], ['MNDTR'], ['EGEPO'], ['NATEN'], ['NTGAZ'], ['NTHOL'], ['NETAS'], ['NIBAS'], ['NUHCM'], ['NUGYO'], ['NRHOL'], ['NRLIN'], ['NURVK'], ['NRBNK'], ['OBAMS'], ['OBASE'], ['ODAS'], ['ODINE'], ['OFSYM'], ['ONCSM'], ['ONRYT'], ['OPET'], ['OPTMA'], ['ORCAY'], ['ORFIN'], ['ORGE'], ['ORMA'], ['OSMEN'], ['OSTIM'], ['OTKAR'], ['OTOKC'], ['OTOSR'], ['OTTO'], ['OYAKC'], ['OYYAT'], ['OYAYO'], ['OYLUM'], ['OZKGY'], ['OZGYO'], ['OZRDN'], ['OZSUB'], ['OZYSR'],
            ['PAMEL'], ['PNLSN'], ['PAGYO'], ['PAPIL'], ['PRFFK'], ['PRDGS'], ['PRKME'], ['PARSN'], ['PBTR'], ['PATEK'], ['PASEU'], ['PSGYO'], ['PCILT'], ['PGSUS'], ['PEKGY'], ['PENGD'], ['PENTA'], ['PEHOL'], ['PSDTC'], ['PETKM'], ['PKENT'], ['PETUN'], ['PINSU'], ['PNSUT'], ['PKART'], ['PLTUR'], ['POLHO'], ['POLTK'], ['PRZMA'], ['QYHOL'], ['QNBFF'], ['QNBFL'], ['QNBVK'], ['QNBFI'], ['QNBFB'], ['QUAGR'], ['QUFIN'], ['RNPOL'], ['RALYH'], ['RAYSG'], ['REEDR'], ['RYGYO'], ['RYSAS'], ['RODRG'], ['ROYAL'], ['RGYAS'], ['RTALB'], ['RUBNS'], ['SAFKR'], ['SANEL'], ['SNICA'], ['SANFM'], ['SANKO'], ['SAMAT'], ['SARKY'], ['SARTN'], ['SASA'], ['SAYAS'], ['SDTTR'], ['SEKUR'], ['SELEC'], ['SELGD'], ['SELVA'], ['SNKRN'], ['SRVGY'], ['KHSTR'], ['SEYKM'], ['SHTRP'], ['SILVR'], ['SNGYO'], ['SKYLP'], ['SMRTG'], ['SMART'], ['SODSN'], ['SOKE'], ['SKTAS'], ['SONME'], ['SNPAM'], ['SUMAS'], ['SUNTK'], ['SURGY'], ['SUWEN'], ['SMRFA'], ['SMRVA'], ['SEKFA'], ['SEKFK'], ['SEGYO'], ['SKYMD'], ['SKBNK'], ['SOKM'], ['DRPHN'],
            ['TOKI'], ['TABGD'], ['TCRYT'], ['TAMFA'], ['TNZTP'], ['TARKM'], ['TATGD'], ['TATEN'], ['TAVHL'], ['TEBFA'], ['TEBCE'], ['TEKTU'], ['TKFEN'], ['TKNSA'], ['TMPOL'], ['TERA'], ['TETMT'], ['TFNVK'], ['TGSAS'], ['TRYKI'], ['TOASO'], ['TRGYO'], ['TLMAN'], ['TSPOR'], ['TDGYO'], ['TRMEN'], ['TSGYO'], ['TUCLK'], ['TUKAS'], ['TRCAS'], ['TUREX'], ['MARBL'], ['TRKFN'], ['TRILC'], ['FNCLL'], ['TCELL'], ['TRKSH'], ['TRKNT'], ['TMSN'], ['TUPRS'], ['THYAO'], ['PRKAB'], ['TTKOM'], ['TTRAK'], ['TBORG'], ['TURGG'], ['GARAN'], ['HALKB'], ['EXIMB'], ['ISATR'], ['ISBTR'], ['ISCTR'], ['ISKUR'], ['KLNMA'], ['TSKB'], ['TURSG'], ['SISE'], ['VAKBN'], ['UFUK'], ['ULAS'], ['ULUFA'], ['ULUSE'], ['ULUUN'], ['UMPAS'], ['USAK'], ['UZERB'], ['ULKER'], ['UNLU'], ['VAKFN'], ['VKGYO'], ['VKFYO'], ['VAKVK'], ['VAKKO'], ['VANGD'], ['VBTYZ'], ['VDFLO'], ['VRGYO'], ['VERUS'], ['VERTU'], ['VESBE'], ['VESTL'], ['VKING'], ['VDFAS'],
            ['YKFKT'], ['YKFIN'], ['YKYAT'], ['YKBNK'], ['YAPRK'], ['YATAS'], ['YFMEN'], ['YATVK'], ['YYLGD'], ['YAYLA'], ['YGGYO'], ['YEOTK'], ['YGYO'], ['YYAPI'], ['YESIL'], ['YBTAS'], ['YIGIT'], ['YONGA'], ['YKSLN'], ['YUNSA'], ['ZEDUR'], ['ZRGYO'], ['ZKBVK'], ['ZKBVR'], ['ZOREN'], ['ZORLF']]
liste_boyut = len(HisseListe)
print(f"Liste Eleman Sayısı: {liste_boyut}")
HisseListe=[['BINHO'], ['AVOD'], ['A1CAP'], ['ACSEL']]
for hisse_sembolu in HisseListe:
    stock_symbol = hisse_sembolu[0]  # Assuming the first element is the stock symbol
    scrape_to_excel(stock_symbol)"""
# Kullanıcıdan hisse senedi sembolünü alın.
stock_symbol = "ASTOR"  # veya input("Lütfen hisse senedi sembolünü giriniz: ")

# Fonksiyonu çağırın.
scrape_to_excel(stock_symbol)