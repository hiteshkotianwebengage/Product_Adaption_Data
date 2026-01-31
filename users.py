'''
    Things to change when we switch between India, GLobal, KSA
    1) Below Client inside init_google_sheet function change the sheet name
    2) In the step 0 below the Driver varibale change between driver.get 
    3) In the step 2 while direct url for publisher change between driver.get
    4) In the step E we have urls to directly go to publisher after data is fetching
    5) This is pretty main we have to change the LC based on the global, india, ksa
'''
import os
import time
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
import gspread
from google.oauth2.service_account import Credentials
from selenium.webdriver.common.action_chains import ActionChains

# =======================
# REGION CONFIG
# =======================
REGION = "KSA"   # options: "INDIA", "GLOBAL", "KSA"

REGION_CONFIG = {
    "INDIA": {
        "base_url": "https://dashboard.in.webengage.com",
        "sheet_name": "User India",
        "publisher_url": "https://dashboard.in.webengage.com/admin/publisher.html?action=list",
        "license_codes": [
            "in~d3a49ca9","in~~9919912c","in~311c47dd","in~14507c7a6","in~~c2ab3610","in~~991990bc","in~~13410618d","in~~99199192","in~~134106156","in~311c4863","in~~15ba205db","in~311c4838","in~11b5642c7","in~~2024c18c","in~76aa307","in~8261735b","in~58adcc15","in~aa131896","in~826173c2","in~14507c7cd","in~58adcd07","in~~10a5cba6d","in~58adcbdb","in~~2024c1a8","in~~47b66667","in~14507c80b","in~76aa34a","in~14507c789","in~14507c784","in~11b564357","in~~c2ab364c","in~~c2ab3662","in~76aa392","in~~10a5cba77","in~8261729b","in~~10a5cbb14","in~14507c76b","in~311c4774","in~~47b66639","in~14507c77b","in~~47b66647","in~58adcc4a","in~~47b665d8","in~~2024c179","in~58adcc70","in~58adcc59","in~~c2ab363b","in~76aa2c5","in~~10a5cba63","in~~9919913a","in~82617341","in~11b5642aa","in~11b56430c","in~~1341061ba","in~~2024c1cd","in~~10a5cbb09","in~~2024c156","in~~134106180","in~aa131782","in~14507c7c3","in~d3a49c7d","in~11b5642a4","in~~c2ab3517","in~76aa2c9","in~~71680ad0","in~~2024c1c6","in~~47b66678","in~~15ba2065c","in~~99199151","in~~99199168","in~~c2ab361a","in~14507c838","in~~1341061ac","in~58adcc11","in~76aa35b","in~~2024c081","in~~134106132","in~~c2ab35d2","in~58adcc57","in~~13410619c","in~~71680aa0","in~d3a49cb0","in~58adcc36","in~aa1317cc","in~~10a5cbb29","in~~15ba205d1","in~~1341061b6","in~~1341061b5","in~~10a5cbb1d","in~d3a49c80","in~11b5642a5","in~~10a5cbac3","in~58adcbd2","in~58adcbda","in~~c2ab3690","in~~47b66670","in~~2024c1bc","in~~c2ab3695","in~~991991c8","in~~15ba20633","in~~47b6663c","in~~1341061cb","in~76aa2a6","in~d3a49c1a","in~~47b66699","in~311c4742","in~~99199207","in~311c472d","in~311c474b","in~~134106200","in~311c4724","in~~134106208","in~~991991c4","in~~134106220","in~~c2ab3675","in~311c488b","in~11b564274","in~~10a5cbb2d","in~~71680b39","in~8261728c","in~14507c728","in~~15ba20670","in~~15ba206a3","in~~15ba206bb","in~~15ba206a9","in~aa1316d4","in~~15ba2068a","in~8261726b","in~~2024c233","in~76aa24d","in~~47b666d5","in~76aa1d8","in~8261722b","in~76aa247","in~~2024c254","in~~99199258","in~311c4708","in~~c2ab35bb","in~~c2ab368c","in~d3a49b8c","in~311c4773","in~~71680b78","in~311c4703","in~~71680b65","in~311c4744","in~11b5642a0","in~~2024c207","in~76aa2b4","in~~10a5cbb42","in~~c2ab36a2","in~~99199206","in~11b564260","in~14507c6ca","in~~71680a90","in~~99199205","in~aa1318ab","in~~47b66709","in~~134106257","in~~c2ab3714","in~d3a49b66","in~aa1316dd","in~58adcb8b","in~~47b666c4","in~d3a49b94","in~d3a49bac","in~aa1316cc","in~~71680b93","in~~10a5cbb66","in~~c2ab36d5","in~aa1316c3","in~~15ba206c2","in~76aa221","in~~c2ab3671","in~~47b66716","in~76aa23c","in~82617226","in~82617217","in~~c2ab36d4","in~8261723d","in~~2024c255","in~d3a49ba1","in~d3a49b86","in~~15ba20658","in~~2024c247","in~14507c71d","in~~991991d0","in~~10a5cbb34","in~76aa273","in~~15ba20652","in~~99199217","in~76aa298","in~~1341061c6","in~~1341061c8","in~aa1316aa","in~~10a5cba3a","in~~10a5cbad8","in~~10a5cbba6","in~~10a5cbbb5","in~76aa21c","in~~c2ab36d9","in~14507c738","in~~15ba206d5","in~14507c6b7","in~d3a49b75","in~11b5641db","in~58adcb61","in~~15ba20690","in~76aa1c5","in~~134106263","in~58adcb79","in~aa1316d1","in~~15ba20672","in~~c2ab3721","in~~71680bbb","in~82617205","in~14507c695","in~11b56420d","in~82617246","in~58adcb94","in~82617203","in~~134106115","in~58adcc83","in~d3a49b80","in~~99199283","in~76aa200","in~~134106273","in~~71680c0c","in~826171d8","in~~1341061b2","in~311c46c9","in~11b5641d0","in~~2024c239","in~~99199233","in~76aa1d3","in~d3a49b43","in~58adcb4b","in~~134106259","in~~10a5cbba9","in~76aa1b3","in~826171d3","in~~10a5cbb79","in~~10a5cbbc9","in~~134106294","in~11b5641ba","in~~47b66722","in~~2024c242","in~d3a49b5a","in~826171c3","in~58adcb40","in~d3a49b3a","in~~2024c276","in~~47b66714","in~~71680ba6","in~~10a5cbbdd","in~~99199068","in~~71680bad","in~~13410624c","in~~10a5cbb76","in~~134106267","in~14507c68c","in~14507c6a9","in~~c2ab3708","in~76aa22a","in~76aa1d0","in~~10a5cbc11","in~~10a5cbc11","in~~10a5cbba4","in~~47b666dc","in~311c4671","in~58adcb30","in~311c467b","in~~47b6675b","in~~991992ab","in~~134106286","in~d3a49c4c","in~~99199277","in~~10a5cbb38","in~~71680c14","in~~13410629c","in~~15ba20749","in~~10a5cbb14","in~~71680bd5","in~311c467c","in~~2024c2aa","in~~71680b12","in~~47b66689","in~~2024c1d7","in~~2024c218","in~311c474a","in~~71680b3c","in~~47b6668a","in~~99199244","in~~2024c246","in~11b564276","in~~134106213","in~~15ba206a8","in~~10a5cbb61","in~11b564256","in~~c2ab36ad","in~76aa268","in~~9919921b","in~~134106216","in~~71680b69","in~~71680b92","in~76aa245","in~311c46d4","in~311c46d3","in~58adcb85","in~~2024c249","in~11b5641d1","in~~47b66730","in~aa131685","in~826172a7","in~76aa1c0","in~311c4665","in~11b564332","in~14507c7d2","in~11b564340","in~aa131666","in~11b5641a6","in~~71680c19","in~~134106253","in~~15ba20752","in~~2024c085","in~d3a49b5d","in~76aa1d7","in~~10a5cbc14","in~aa131655","in~~47b66750","in~76aa241","in~~2024c2a0","in~14507c681","in~aa131667","in~~134106266","in~11b5641a0","in~~15ba20741","in~~991992a7","in~~991992b1","in~~2024c2c1","in~~71680b61","in~76aa206","in~~10a5cbc2c","in~826171b0","in~~71680c29","in~aa131676","in~~71680bb9","in~d3a49b24","in~~c2ab3735","in~aa131652","in~14507c67b","in~aa131675","in~14507c65b","in~11b5641a9","in~~2024c27c","in~11b5641aa","in~d3a49b10","in~aa13163d","in~~15ba20753","in~d3a49b0b","in~d3a49b14","in~~991992c6","in~~15ba2076c","in~~71680c2b","in~~1341062c9","in~14507c647","in~82617199","in~~71680c38","in~58adcb50","in~~991992aa","in~76aa201","in~58adcb08","in~~991992c7","in~~47b66733","in~~10a5cbc25","in~aa131650","in~aa13163a","in~11b56418d","in~11b564191","in~~2024c2b8","in~311c4663","in~76aa1a2","in~~15ba2074d","in~~c2ab3781","in~~1341062bb","in~~991992c4","in~~10a5cbc2d","in~~1341062c1","in~~99199240","in~~71680b76","in~~991992cc","in~~47b6678c","in~311c4664","in~14507c641","in~~71680c30","in~aa13164b","in~~991992a4","in~~15ba20759","in~~15ba205c0","in~~2024c231","in~76aa1ac","in~11b5641b1","in~~47b6677d","in~58adcb36","in~aa13166b","in~~991992d1","in~~1341062c2","in~~99199081","in~14507c63b","in~~c2ab3786","in~11b564246","in~~99199278","in~11b56417b","in~aa131665","in~~71680b90","in~14507c666","in~aa131632","in~76aa20d","in~311c464b","in~311c4766","in~~c2ab3761","in~~71680c4c","in~11b564177","in~11b564172","in~d3a49ad8","in~~47b66782","in~11b56417c","in~11b564181","in~~c2ab3789","in~11b5642d2","in~311c4646"
        ]
    },
    "GLOBAL": {
        "base_url": "https://dashboard.webengage.com",
        "sheet_name": "User Global",
        "publisher_url": "https://dashboard.webengage.com/admin/publisher.html?action=list",
        "license_codes": [
            "~47b6574d","d3a4ab38","~47b66864","76aac96","8261829c","~10a5cad0b","58add307","~10a5cabbb","58add216","~10a5cb2a9","~2024b5d8","d3a4a457","311c5625","~134105a60","58add423","311c5642","14507cd00","d3a4a301","~15ba20153","d3a4ac1c","~1341059b6","~10a5cab6a","14507cd51","~15ba1d846","~10a5cac40","~7167db84","14507cc77","58add283","~99198968","~15ba1da68","14507cda4","58add2d9","~134105a52","~7167db54","~47b66064","76aa813","~716800b0","aa13266b","826174d0","~47b65b6c","82617c25","~47b6607a","~15ba2020b","~76aa7a9","82617894","82617822","76aa800","~c2ab3242","~15ba1d691","~10a5cb63c","~c2ab3108","~134105a8c","76ab0a5","~c2ab2dd0","~15ba20105","82617757","aa131c59","311c4b69","~2024bb2d","~134105a04","14507cd4d","~13410604b","76aa78b","~71680543","~2024b99a","~71680588","~9919839a","aa131c84","76aa858","~134105251","~99198a20","~1341056a0","~2024b6c6","~15ba20116","76aa833","~991983db","76aa7a3","14507cba6","14507d14d","~71680655","82617754","~2024bb10","~c2ab2851","~99198b61","~10a5cb5d1","~10a5cb677","~10a5cb557","~c2ab313b","~134105965","~1341059c5","58adc5c7","~7168057d","~99198ab8","~15ba1ddc6","~47b65848","d3a4a32d","311c4c4b","~47b66045","~15ba201a6","76aa124","~9919868d","76a9c30","~311c4b76","58adca91","76aa76b","82617b34","~2024bada","d3a4a403","8261827a","~71680577","~99198a29","~2024bad5","~2024b291","58add69d","~134105b84","d3a4a6dd","311c558d","~c2ab3033","~15ba1ddc2","82618089","58add346","~1341061bb","d3a4a69c","58add2da","~716805d8","~15ba20214","~10a5cb6b0","~1341056bb","aa132703","~aa1321c5","~c2ab2c0c","~9919871c","~47b66614","~991981d3","14507cc74","8261786b","11b564b69","aa132225","311c4c14","311c4c11","76aa762","311c4bbb","14507cba8","58add7aa","11b564830","76aac69","76aab88","~71680627","~15ba20042","~oldetmoney","11b564720","~c2ab275a","~c2ab2bc0","~15ba1d70a","~c2ab2ba2","8261812c","~47b6665c","~1341059cb","~991989d1","~old2024c085","d3a4a36a","~10a5cac99","~oldmagma1","~oldmagma2","~oldmagma3","11b56527d","11b5646ca","~c2ab2c08","~11b5646b8","~d3a4a286","76aa7c6","76aa85d","~2024ba8b","~47b65b94","~7168053d","~c2ab3083","58addc40","76aa844","~d3a49c4c-old","76aa868","d3a4a663","~15ba2019a","~7168069d","~oldrangde","~10a5cb533","~10a5cb636","~c2ab26b8","~99198226","~2024bbb6","~99198a14","82617775","~134105aac","11b56470b","aa13264a","~c2ab3042","~99198a18","11b5650c0","311c4bc4","d3a4a420","~10a5cb20c","14507cc0a","~15ba1db98","~14507ccb9","~134105a45","~15ba200d7","~11b564836","~14507ccc0","aa131c70","~47b661c8","~7168071b","~10a5cb278","~2024bb90","~82617869","~oldUPES","14507cba1","d3a4a72a","~71680632","~47b65a1c","82617779","~10a5cb53d","311c4d24","58add3a2","~47b660d4","~134105732","~9919837c","~311c4dc3","~2024bb26","~76aab32","~c2ab323a","old~2024c1a3","826182a0","~10a5cb24b","11b56488a","~15ba20234","~15ba2063d","~47b661ab","~47b65875","~99198624","14507cd97","~134105353","d3a4ab04","~ c2ab260c","~47b66257","~c2ab3091","~c2ab30aa","~15ba20080","d3a4a3b7","~aa131bd4","~134105b82","~134105a36","58add667","~47b65c2a","311c4b7c","aa131752","11b564bc3","82617bba","~716802a1","~14507d169","~d3a4a64d","~aa1321a1","~716802d0","~7168026d","~15ba1db9b","~2024b725","~2024b72c","~7168028a","~47b65bd8","~99199107","~oldid","~2024b7a2","aa132182","~10a5cb319","~15ba1dbd6","~10a5cb2a0","~1341056c2","~2024b80a","aa13210d","~11b564252","~47b65c27","14507d197","14507d153","~aa131717","311c5166","~15ba1dc69","~10a5cb33d","~c2ab2c5d","~82616dd3","~14507c905","~58adc916","~11b563c29","~311c4664","~c2ab3c28","~7167db9d","~99198b2b","~oldedelwiess","~2024b1c1","~10a5cb621","~15ba20147","~15ba1dc2d","~10a5cb283","82617aac","old~10a5cbb38","58addb9c","58adcba4","~oldsasai","~ c2ab36a7","~9919922b","~15ba1dd60","~9919879c","~15ba1dda1","~716803a0","~10a5cb41b","82617a35","~71680426","~15ba206bc","~teamGreatLearning","old~7167d2da","~10a5cb41c","d3a4b631","~c2ab3713","d3a4a554","old~76aab14","58add639","~2024b6b1","~716802c1","~c2ab2d61","~47b65736","aa132108","~13410583d","~716806c3","58add571","old~1341061b2","~2024b742","~15ba2009b","~2024b7a6","~47b65ca0","d3a4a5c9","~82617225","76abb05","~15ba1cc68","~134105ba0","82617a1a","~2024b8a7","~82617957","76aab70","d3a49b72","~10a5cac62","~991988b2","14507d028","~10a5cb41c","d3a4a4aa","311c5128","~oldaccount","82618240","76abad2","~47b66522","~15ba1cc70","~15ba20218","~7167d2da","~47b66119","~716805b9","14507d167","76aa239","14507d0c5","11b565a86","aa133168","~15ba201dc","~1341047d4","~716806db","~47b65c58","~10a5cb29d","~10a5cbb86","aa1318a4","~15ba20712","~2024a7ad","aa133154","d3a4a4c6","311c561d","~716801dd","~58add514","~2024b7a3","~71680374","~9919835a","~82617a45","d3a4a292","~9919779b","311c6080","~7167d365","~9919921c","~716806cb","76aab70","~47b666aa","~2024a7c9","~134104786","~71680bc5","76aa235","~2024b77c","~1341047d6","311c6141","d3a4b607","11b565a2a","82618a47","~10a5ca344","11b565a4b","~76aa276","~13410606b","~2024c07d","58b005a5","76ab982","58add2d9","14507d153","~47b66726","1450800cc","82618a78","76aba82","~99197808","~9919786b","76ab099","~99197879","~c2ab1d22","~134105365","~7167dd04","aa133088","~15ba20518","~11b565a1","~c2ab1d64","~99197854","d3a4a339","~c2ab3737","~134105770","~82618284","11b564400","58b005b2","76aab2b","311c6114","~15ba1cd7b","~15ba1cd57","8261895b","~2024a872","~47b64d14","~145080099","~2024a898","~10a5cb63c","~2024bb10","~99197908","~10a5cbbc3","311c6137","~145080038","d3a4b4bc","11b56595a","~10a5cbb7a","d3a4a3d2","311c617c","~1341058da","~15ba1dd60","~15ba1cd06","311c6067","~10a5cb5d1","~7167d43b","11b56597b","76ab9b6","76ab933","76aa777","76aab14","~15ba1cc97","145080012","~47b64db3","76ab966","~134104915","~9919792b","76ab97b","d3a4b5a8","145080010","~2024a887","145080030","~47b64d1c","~47b64dc6","11b565a3b","d3a4b4a3","~2024ba16","~c2ab3054","~99197905","~10a5cb448","~2024a920","~c2ab1db9","311c6015","76ab975","~2024a932","~99197925","~7167d473","~10a5ca45d","d3a4b49c","58b00497","~2024a938","d3a4b48c","~7167d481","~amplicom","~15ba1cda8","82618947.0","~TapsiMarket","58b004a0","~15ba1cdc0","~134104929","~GreenlifeEwaste","311c5018","~15ba1cd29","76ab953","~7167d478","11b565ab8","d3a4b501","311c6036","aa133008","~2024a928","76ab919","~15ba1cdc7","311c6106","~2024a8d9","~7167d459","~47b64dcd","76ab94b","~10a5ca413","~c2ab1d61","~47b64d69","11b565933","aa13306d","14507ddc0","76ab948","~99197942","311c601a","311c5ddd","~15ba1cd71","11b56593a","14507ddd3","58add3d1","~c2ab3a30","58b00572","d3a4b47a","58b00482","~10a5ca49a","58b0047d","aa131b8b","~134104910","~10a5cb607","~7167d4a7","311c50c2","76ab96b","~13410487b","311c50ac","~99198826","76aab90","311c4c52","~71680307","~15ba1dcd4","~47b64d09","~99198b33","~c2ab1d11","d3a4a33c","~c2ab1dda","aa132dcb","76ab943","58b00487","11b565931","~134104942","aa132dc1","~15ba1d007","~134104952","76ab929","~1341053c6","d3a4b479","aa132dac","~c2ab1dc3","11b565915","76ab940","58add629","~47b64c57","~2024a96d","~7167d309","311c6107","~c2ab2d96","aa131c6a","~7167d4b3","~7168061a","76ab912","~2024a973","311c5dac","~134104968","~2024a7c4","58b00457","76ab922","~7167d4b4","~10a5ca4a2"
        ]
    },
    "KSA": {
        "base_url": "https://dashboard.ksa.webengage.com",
        "sheet_name": "test1",
        "publisher_url": "https://dashboard.ksa.webengage.com/admin/publisher.html?action=list",
        "license_codes": [
            "ksa~~15ba20526","ksa~82617412","ksa~~716809b4","ksa~14507c890","ksa~~2024c070","ksa~76aa41c","ksa~~716809bd","ksa~11b564409","ksa~~47b6652a","ksa~aa1318a1","ksa~~47b6652c","ksa~58adcd4c","ksa~11b564406","ksa~~15ba2051c","ksa~aa131897","ksa~~134106071","ksa~82617408","ksa~58adcd55","ksa~~15ba20523","ksa~~10a5cb9bd","ksa~d3a49d49","ksa~11b5643db","ksa~d3a49d4a","ksa~d3a49d44","ksa~58adcd54","ksa~aa13189b","ksa~311c489a","ksa~82617404","ksa~~2024c08a","ksa~~134106076","ksa~14507c891","ksa~~716809c9","ksa~d3a49d46","ksa~~47b66537","ksa~~134106074","ksa~~2024c07d","ksa~~2024c085","ksa~82617402","ksa~~13410607a","ksa~11b564403","ksa~~716809ba","ksa~~10a5cb9c4","ksa~~99199078","ksa~~15ba20518","ksa~~134106069","ksa~311c4892","ksa~~99199083","ksa~aa1318a0","ksa~~13410606b","ksa~11b5643d5","ksa~~134106080","ksa~58adcd47","ksa~58adcd44","ksa~11b5643d3","ksa~826173db","ksa~d3a49d41","ksa~~134106084","ksa~~99199073","ksa~~99199087","ksa~aa131893","ksa~~15ba20531","ksa~76aa3da","ksa~~2024c091","ksa~826173dc","ksa~82617401","ksa~aa131890","ksa~14507c89c","ksa~~10a5cb9d1","ksa~~2024c08d","ksa~~47b66522"
        ]
    }
}

def init_google_sheet():
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = Credentials.from_service_account_file(
        "/Users/admin/Desktop/Python Script/agreement_file_pasting/mycred-googlesheet.json",
        scopes=scopes
    )
    client = gspread.authorize(creds)

    sheet_name = REGION_CONFIG[REGION]["sheet_name"]
    sheet = client.open_by_key(
        "1sathL7caATX3PnV2urKhp8UpCLOcwdTIxzkJkA4Kar4"
    ).worksheet(sheet_name)

    return sheet

try:
    sheet = init_google_sheet()
    print("✅ Connected to Google Sheets")
except Exception as e:
    print(f"❌ Failed to connect to Google Sheets: {e}")
    exit()


# ---------- STEP 0: SETUP PERSISTENT PROFILE ----------
script_dir = os.path.dirname(os.path.abspath(__file__))
user_data_dir = os.path.join(script_dir, "selenium_profile")

options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
options.add_argument(f"--user-data-dir={user_data_dir}")

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
wait = WebDriverWait(driver, 120) 

# Step 1: Open admin dashboard
BASE_URL = REGION_CONFIG[REGION]["base_url"]
driver.get(f"{BASE_URL}/admin")

# --- GIVE IT A MOMENT TO REDIRECT ---
print("⏳ Waiting for page to settle...")
time.sleep(3) 

# ---------- STEP 2: SMART PAGE DETECTION ----------
print(f"🔍 Current URL: {driver.current_url}")

# Give it one refresh if we aren't where we expect to be
if "publisher.html" not in driver.current_url:
    print("🔄 Refreshing to ensure state...")
    driver.refresh()
    time.sleep(2)

if "publisher.html" in driver.current_url:
    print("✅ Already on Publishers page via Cookies.")

else:
    try:
        print("🔍 Checking if Publishers link is visible...")
        # Try a more generic XPath that finds the link even if text is weird
        publisher_xpath = "//a[contains(@href, 'publisher.html')]"
        wait_short = WebDriverWait(driver, 15) # Increased to 15
        
        link = wait_short.until(EC.element_to_be_clickable((By.XPATH, publisher_xpath)))
        link.click()
        print("✅ Clicked Publishers from sidebar.")

    except Exception:
        print("📍 Sidebar link not found. Starting Deep Navigation...")

        # FIX: The profile head might be nested. Let's use a simpler selector.
        try:
            profile_xpath = "//div[contains(@class,'pop-over__head')] | //div[contains(@class,'noselect')]"
            dropdown = wait.until(EC.element_to_be_clickable((By.XPATH, profile_xpath)))
            driver.execute_script("arguments[0].click();", dropdown) # JS click is safer here
            print("✅ Opened Profile Dropdown")
            time.sleep(1.5)

            # Click Super Admin - Use a more flexible text match
            super_admin_xpath = "//a[normalize-space()='Super Admin' or contains(text(),'Super Admin')]"
            sa_btn = wait.until(EC.element_to_be_clickable((By.XPATH, super_admin_xpath)))
            sa_btn.click()
            print("✅ Clicked Super Admin")

            # Final check for Publisher link after landing in Super Admin
            final_publisher_xpath = "//a[contains(@href,'publisher.html')]"
            wait.until(EC.element_to_be_clickable((By.XPATH, final_publisher_xpath))).click()
        except Exception as e:
            print(f"❌ Deep Navigation failed: {e}")
            driver.save_screenshot("nav_failure.png")
            # If everything fails, try going to the URL directly as a last resort
            print("🚀 Attempting direct URL navigation...")
            PUBLISHER_URL = REGION_CONFIG[REGION]["publisher_url"]
            driver.get(PUBLISHER_URL)


print("🎯 SUCCESS: You are now on the Publishers page.")

def search_by_license(driver, wait, license_code):
    license_input = wait.until(
        EC.presence_of_element_located((By.NAME, "licenseCode"))
    )
    license_input.clear()
    license_input.send_keys(license_code)

    search_btn = wait.until(
        EC.element_to_be_clickable(
            (By.XPATH, "//button[normalize-space()='Search']")
        )
    )
    search_btn.click()
    print("✅ Succesfully entered the LC and clicked search button")

def open_actions_dropdown(driver, wait, license_code):
    print(f"⏳ Opening Actions dropdown for {license_code}...")
    
    # Improved XPath: Finds the specific TR containing the LC, then the toggle button inside it
    toggle_xpath = f"//tr[contains(., '{license_code}')]//button[contains(@class, 'dropdown-toggle')]"
    
    try:
        dropdown_btn = wait.until(EC.element_to_be_clickable((By.XPATH, toggle_xpath)))
        # Force scroll and use JS click to bypass any overlapping UI elements
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", dropdown_btn)
        time.sleep(0.5)
        driver.execute_script("arguments[0].click();", dropdown_btn)
        print("✅ Successfully clicked the specific dropdown toggle")
    except Exception as e:
        print(f"⚠️ Could not click dropdown for {license_code}: {e}")
        # If it fails, we try a refresh as a fallback
        driver.refresh()
        time.sleep(2)

def click_request_access(wait):
    request_access = wait.until(
        EC.element_to_be_clickable(
            (By.XPATH, "//a[contains(@class,'requestAccess')]")
        )
    )
    request_access.click()
    print("✅ Succesfully opened the access dialog box")

def handle_request_modal(wait, driver):
    print("⏳ Handling request modal...")
    
    # Give the modal a second to fully animate and render
    time.sleep(2)

    # 1. SMART IFRAME DETECTION
    # We only switch if an iframe actually exists AND is visible
    iframes = driver.find_elements(By.TAG_NAME, "iframe")
    switched = False
    for frame in iframes:
        if frame.is_displayed():
            driver.switch_to.frame(frame)
            switched = True
            print("↔️ Switched to active iframe")
            break

    try:
        # 2. WAIT FOR DROPDOWN
        # We use a very broad XPath to find the select by its ID or Name
        dropdown_xpath = "//select[@id='roleIdField' or @name='roleEId']"
        role_dropdown = wait.until(
            EC.presence_of_element_located((By.XPATH, dropdown_xpath))
        )
        
        # 3. FORCE SELECTION VIA JAVASCRIPT
        # This is the most reliable way to select 'Viewer' regardless of UI quirks
        print("🚀 # React-safe role selection (required only during first-time access request)")
        Select(role_dropdown).select_by_visible_text("Viewer")
        
        # 4. FILL COMMENT
        comment_box = wait.until(EC.presence_of_element_located((By.ID, "commentText")))
        comment_box.clear()
        comment_box.send_keys("access request")

        # 5. CLICK REQUEST
        request_btn = driver.find_element(By.XPATH, "//button[contains(.,'Request')]")
        driver.execute_script("arguments[0].click();", request_btn)
        print("✅ Successfully submitted access request")

    except Exception as e:
        print(f"❌ Error inside modal: {e}")
        driver.save_screenshot("error_modal_view.png")
    
    finally:
        if switched:
            driver.switch_to.default_content()
            print("↔️ Switched back to main content")

def close_modal_if_exists(driver):
    """Closes the modal only if it is still visible."""
    try:
        # Check for the close button with a very short timeout
        close_btn = WebDriverWait(driver, 3).until(
            EC.element_to_be_clickable((By.ID, "cboxClose"))
        )
        close_btn.click()
        print("✅ Manually closed the dialog box")
        time.sleep(1)
    except:
        print("ℹ️ Modal already closed or close button not found (which is fine)")

def click_edit(driver, wait, license_code):
    print(f"⏳ Clicking Edit for {license_code}...")
    
    # Find the Edit link inside the row that matches the License Code
    edit_xpath = f"//tr[contains(., '{license_code}')]//a[contains(@href, '/publisher/edit')]"
    
    edit_btn = wait.until(
        EC.element_to_be_clickable((By.XPATH, edit_xpath))
    )
    
    edit_btn.click()
    print("✅ Clicked the specific edit button")


def click_users(wait):
    print("⏳ Clicking Users...")

    users_xpath = (
        "//a[contains(@href,'/users/overview') and .//span[text()='Users']]"
    )

    wait.until(
        EC.element_to_be_clickable((By.XPATH, users_xpath))
    ).click()

    print("✅ Users opened")

def switch_users_delta(wait, driver, mode):
    """ mode = 'WoW' or 'MoM' """
    print(f"🔁 Switching Users delta to {mode}")

    dropdown_head = wait.until(
        EC.element_to_be_clickable(
            (By.XPATH, "//th//div[contains(@class,'pop-over__head')]")
        )
    )

    driver.execute_script("arguments[0].click();", dropdown_head)

    option = wait.until(
        EC.element_to_be_clickable(
            (By.XPATH, f"//span[normalize-space()='{mode}']")
        )
    )

    driver.execute_script("arguments[0].click();", option)

    # 🔹 CRITICAL: wait for DOM refresh
    time.sleep(1.2)
    wait_for_users_table_or_empty(driver)

def extract_users_table(driver, license_code, delta_type):
    rows_data = []

    has_data = wait_for_users_table_or_empty(driver)

    if not has_data:
        print(f"⚠️ No Users data for {delta_type}")
        return [[
            license_code,
            "NO_DATA",
            "NO_DATA",
            "NO_DATA",
            delta_type,
            "NO_DATA"
        ]]

    rows = driver.find_elements(
        By.XPATH,
        "//tbody/tr[contains(@class,'table__row')]"
    )

    for row in rows:
        try:
            cells = row.find_elements(By.XPATH, "./td")
            if len(cells) < 4:
                continue

            channel = cells[0].text.strip() or "UNKNOWN"
            reach_pct = cells[1].text.strip() or "0%"
            reach_count = cells[2].text.strip() or "0"

            # Change %
            change = cells[3].text.strip() or "0.00%"

            rows_data.append([
                license_code,
                channel,
                reach_pct,
                reach_count,
                delta_type,
                change
            ])

        except:
            continue

    return rows_data

def extract_users_reachability(driver, wait, license_code):
    print("📥 Extracting Users Reachability (WoW + MoM)...")

    all_rows = []

    for mode in ["WoW", "MoM"]:
        switch_users_delta(wait, driver, mode)
        rows = extract_users_table(driver, license_code, mode)
        all_rows.extend(rows)

    return all_rows

def wait_for_users_table_or_empty(driver, timeout=6):
    """
    Waits for:
    - users rows
    - footer row (Overall)
    - empty state
    Returns True if data rows exist, False otherwise
    """
    try:
        WebDriverWait(driver, timeout).until(
            lambda d: (
                d.find_elements(By.XPATH, "//tbody/tr[contains(@class,'table__row')]")
                or d.find_elements(By.XPATH, "//tfoot/tr")
                or "No data" in d.page_source
            )
        )
    except:
        return False

    rows = driver.find_elements(By.XPATH, "//tbody/tr[contains(@class,'table__row')]")
    return len(rows) > 0


def append_to_sheet(sheet, rows):
    if rows:
        sheet.append_rows(rows, value_input_option="USER_ENTERED")
        print("📤 Data pushed to Google Sheets")
    else:
        print("⚠️ No data to push")

def log_error_to_sheet(sheet, license_code, stage, error_reason):
    print(f"📝 Logging error for {license_code} at stage: {stage}")

    row = [
        license_code,
        "ERROR",
        stage,
        error_reason[:300],  # keep it readable
        time.strftime("%Y-%m-%d %H:%M:%S")
    ]

    sheet.append_row(row, value_input_option="USER_ENTERED")

LICENSE_CODES = REGION_CONFIG[REGION]["license_codes"]

for code in LICENSE_CODES:
    print(f"\n▶ Processing {code}")
    try:
        # Step A: Search and land on result
        search_by_license(driver, wait, code)
        
        # Step B: Try to get access (only if needed)
        try:
            open_actions_dropdown(driver, wait, code) # Added driver and code
            
            request_btn_xpath = f"//tr[contains(., '{code}')]//a[contains(@class,'requestAccess')]"
            if len(driver.find_elements(By.XPATH, request_btn_xpath)) > 0:
                click_request_access(wait)
                handle_request_modal(wait, driver)
                time.sleep(2)
                close_modal_if_exists(driver)
            else:
                print("ℹ️ Access already available. Moving to Edit.")
                # If dropdown is open but we don't need it, refresh or Esc
                ActionChains(driver).send_keys(Keys.ESCAPE).perform() 
        except Exception as e:
            print(f"⚠️ Access step skipped: {e}")

        # --- Updated Step C: Pass 'code' to the edit function ---
        main_window = driver.current_window_handle
        click_edit(driver, wait, code)

        # SWITCH TO NEW TAB
        for window_handle in driver.window_handles:
            if window_handle != main_window:
                driver.switch_to.window(window_handle)
                break
        print(f"↔️ Switched to Edit tab for {code}")

        # Step D: Extract Data
        try:
            click_users(wait)

            users_rows = extract_users_reachability(driver, wait, code)
            append_to_sheet(sheet, users_rows)

            time.sleep(1)

            users_rows = extract_users_reachability(driver, wait, code)

            append_to_sheet(sheet, users_rows)

        except Exception as e:
            log_error_to_sheet(
                sheet,
                code,
                stage="USERS_REACHABILITY",
                error_reason=str(e)
            )
    
    finally:
        # Step E: Cleanup for next iteration
        # Close extra tabs and go back to the list
        if len(driver.window_handles) > 1:
            driver.close() # Closes current (Edit) tab
            driver.switch_to.window(main_window)
        
        PUBLISHER_URL = REGION_CONFIG[REGION]["publisher_url"]
        driver.get(PUBLISHER_URL)

        time.sleep(2)
