
# <a name="create-a-dictionary-task-pane-add-in"></a>辞書の作業ウィンドウ アドインを作成する


この記事では、Word 2013 のドキュメントでユーザーの現在の選択範囲に対応する辞書の定義や類義語辞典の類義語を表示する作業ウィンドウ アドインの例と、それに付随する Web サービスについて取り上げます。 

辞書の Office アドインは、標準的な作業ウィンドウ アドインを基盤として、辞書の XML Web サービスに対するクエリの機能と、取得した定義を Office アプリケーションの UI 上の別の場所に表示する機能が追加されたものです。 

一般的な辞書作業ウィンドウ アドインの場合、ユーザーが自分のドキュメントで単語または語句を選択すると、アドインの背景にある JavaScript ロジックにより、この選択はディクショナリ プロバイダーの XML Web サービスに渡されます。ディクショナリ プロバイダーの Web ページは更新され、ユーザーに選択範囲の定義が表示されます。XML Web サービス コンポーネントは、OfficeDefinitions XML スキーマが定める形式で、最大 3 つの定義を返します。アプリはこれを、ホストの Office アプリケーションの UI 上の別の場所に表示します。図 1 は、Word 2013 で実行されている Bing ブランドの辞書アドインの選択および表示エクスペリエンスを示しています。

**図 1.選択した語句の定義を表示する辞書アドイン**


![定義が表示されている辞書アプリ](../../images/DictionaryAgave01.jpg)

辞書アドインの HTML UI で 「**さらに表示**」リンクをクリックして、作業ウィンドウ内に詳細情報を表示するか、別のブラウザー ウィンドウを開いて、選択した単語または語句の完全な Web ページを表示するかはユーザー次第です。図 2 は、ユーザーがインストール済みの辞書を簡単に起動するための、**[定義]** コンテキスト メニュー コマンドを示しています。図 3 から 5 は、辞書 XML サービスを使用して Word 2013 で定義を提供する、Office UI の場所を示しています。

**図 2.コンテキスト メニューの定義コマンド**



![コンテキスト メニューの定義](../../images/DictionaryAgave02.jpg)

**図 3.スペル チェック ウィンドウと文章校正ウィンドウでの定義の表示**


![スペル チェック ウィンドウと文章校正ウィンドウでの定義の表示](../../images/DictionaryAgave03.jpg)

**図 4.類義語辞典ウィンドウでの定義の表示**


![類義語辞典ウィンドウでの定義の表示](../../images/DictionaryAgave04.jpg)

**図 5.読み取りモードでの定義**


![読み取りモードでの定義](../../images/DictionaryAgave05.jpg)

辞書の検索機能を持つ作業ウィンドウ アドインを作成するには、次の 2 つの主要なコンポーネントを作成します。 


- XML Web サービス。辞書サービスで定義を検索し、辞書アドインが利用および表示できる XML 形式でその定義を返します。
    
- 作業ウィンドウ アドイン。ユーザーの現在の選択範囲を辞書の Web サービスに送信し、定義を表示します。必要に応じてその値をドキュメントに挿入することもできます。
    
以下のセクションでは、これらのコンポーネントの作成方法の例を示します。

## <a name="creating-a-dictionary-xml-web-service"></a>辞書の XML Web サービスの作成


XML Web サービスでは、クエリを OfficeDefinitions XML スキーマに準拠した XML で Web サービスに返す必要があります。以下の 2 つのセクションでは、OfficeDefinitions XML スキーマについて説明し、この XML 形式でクエリを返す XML Web サービスのコーディング方法の例を示します。


### <a name="officedefinitions-xml-schema"></a>OfficeDefinitions XML スキーマ

次のコードは、OfficeDefinitions XML スキーマの XSD を示します。


```XML
<?xml version="1.0" encoding="utf-8"?>
<xs:schema
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:xs="http://www.w3.org/2001/XMLSchema"
  targetNamespace="http://schemas.microsoft.com/NLG/2011/OfficeDefinitions"
  xmlns="http://schemas.microsoft.com/NLG/2011/OfficeDefinitions">
  <xs:element name="Result">
    <xs:complexType>
      <xs:sequence>
        <xs:element name="SeeMoreURL" type="xs:anyURI"/>
        <xs:element name="Definitions" type="DefinitionListType"/>
      </xs:sequence>
    </xs:complexType>
  </xs:element>
  <xs:complexType name="DefinitionListType">
    <xs:sequence>
      <xs:element name="Definition" maxOccurs="3">
        <xs:simpleType>
          <xs:restriction base="xs:normalizedString">
            <xs:maxLength value="400"/>
          </xs:restriction>
        </xs:simpleType>
      </xs:element>
    </xs:sequence>
  </xs:complexType>
</xs:schema>
```

OfficeDefinitions スキーマに準拠した XML では、ルートの **Result** 要素の中に **Definitions** 要素を 1 個記述し、その子要素として 0 から 3 個の **Definition** 要素に定義を記述して返します。それぞれの定義は最大 400 文字です。また、辞書サイトの結果ページの URL を **SeeMoreURL** 要素で指定する必要があります。次の例は、OfficeDefinitions スキーマに準拠して返す XML の構造を示します。




```XML
<?xml version="1.0" encoding="utf-8"?>
<Result xmlns="http://schemas.microsoft.com/NLG/2011/OfficeDefinitions">
  <SeeMoreURL xmlns="">www.bing.com/dictionary/search?q=example</SeeMoreURL>
  <Definitions xmlns="">
    <Definition>Definition1</Definition>
    <Definition>Definition2</Definition>
    <Definition>Definition3</Definition>
  </Definitions>
 </Result>

```


### <a name="sample-dictionary-xml-web-service"></a>辞書の XML Web サービスのサンプル

次の C# コードは、辞書クエリの結果を OfficeDefinitions XML 形式で返す XML Web サービスのコードの簡単な作成例です。


```C#
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Services;
using System.Xml;
using System.Text;
using System.IO;
using System.Net;

/// <summary>
/// Summary description for _Default
/// </summary>
[WebService(Namespace = "http://tempuri.org/")]
[WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
// To allow this web service to be called from script, using ASP.NET AJAX, uncomment the following line. 
// [System.Web.Script.Services.ScriptService]
public class WebService : System.Web.Services.WebService {

    public WebService () {

        // Uncomment the following line if using designed components 
        // InitializeComponent(); 
    }

    // You can replace this method entirely with your own method that gets definitions
    // from your data source, and then formats it into the OfficeDefinitions XML format. 
    // If you need a reference for constructing the returned XML, you can use this example as a basis.
    [WebMethod]
    public XmlDocument Define(string word)
    {

        StringBuilder sb = new StringBuilder();
        XmlWriter writer = XmlWriter.Create(sb);
        {
            writer.WriteStartDocument();
            
                writer.WriteStartElement("Result", "http://schemas.microsoft.com/NLG/2011/OfficeDefinitions");

            // See More URL should be changed to the dictionary publisher's page for that word on their website.
                    writer.WriteElementString("SeeMoreURL", "http://www.bing.com/search?q=" + word);

                    writer.WriteStartElement("Definitions");
            
                        writer.WriteElementString("Definition", "Definition 1 of " + word);
                        writer.WriteElementString("Definition", "Definition 2 of " + word);
                        writer.WriteElementString("Definition", "Definition 3 of " + word);
                   
                    writer.WriteEndElement();


                writer.WriteEndElement();
            
            writer.WriteEndDocument();
        }
        writer.Close();

        XmlDocument doc = new XmlDocument();
        doc.LoadXml(sb.ToString());

        return doc;
    }
   

}
```


## <a name="creating-the-components-of-a-dictionary-add-in"></a>辞書アドインのコンポーネントの作成


辞書アドインは 3 つの主要なコンポーネント ファイルで構成されます。


- アドインについての情報を記述した XML マニフェスト ファイル
    
- アドインの UI を記述した HTML ファイル
    
- ユーザーの選択範囲をドキュメントから取得し、選択範囲をクエリとして Web サービスに送信し、返された結果をアドインの UI に表示するロジックを記述した JavaScript ファイル。
    

### <a name="creating-a-dictionary-add-ins-manifest-file"></a>辞書アドインのマニフェスト ファイルの作成

辞書アドインのマニフェスト ファイルの例を次に示します。


```XML
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.0" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="TaskPaneApp">
  <Id>7164e750-dc86-49c0-b548-1bac57abdc7c</Id>
  <Version>15.0</Version>
  <ProviderName>Microsoft Office Demo Dictionary</ProviderName>
  <DefaultLocale>en-us</DefaultLocale>
  <!--DisplayName is the name that will appear in the user's list of applications.-->
  <DisplayName DefaultValue="Microsoft Office Demo Dictionary" />
  <!--Description is a 2-3 sentence description of this dictionary. -->
  <Description DefaultValue="The Microsoft Office Demo Dictionary is an example built to demonstrate how a publisher could create a dictionary that integrates with Office. It does not return real definitions." />
  <!--IconUrl is the URI for the icon that will appear in the user's list of applications.-->
  <IconUrl DefaultValue="http://officeimg.vo.msecnd.net/_layouts/images/general/office_logo.jpg" />
  <!--Capabilities specifies the kind of host application your dictionary add-in will support. You shouldn't have to modify this area.-->
  <Capabilities>
    <Capability Name="Workbook"/>
    <Capability Name="Document"/>
    <Capability Name="Project"/>
  </Capabilities>
  <DefaultSettings>
    <!--SourceLocation is the URL for your dictionary-->
    <SourceLocation DefaultValue="http://christophernlg/ExampleDictionary/DictionaryHome.html" />
  </DefaultSettings>
  <!--Permissions is the set of permissions a user will have to give your dictionary. If you need write access, such as to allow a user to replace the highlighted word with a synonym, use ReadWriteDocument. -->
  <Permissions>ReadDocument</Permissions>
  <Dictionary>
    <!--TargetDialects is the set of regional languages your dictionary contains. For example, if your dictionary applies to Spanish (Mexico) and Spanish (Peru), but not Spanish (Spain), you can specify that here. Do not put more than one language (for example, Spanish and English) here. Publish separate languages as separate dictionaries. -->
    <TargetDialects>
      <TargetDialect>EN-AU</TargetDialect>
      <TargetDialect>EN-BZ</TargetDialect>
      <TargetDialect>EN-CA</TargetDialect>
      <TargetDialect>EN-029</TargetDialect>
      <TargetDialect>EN-HK</TargetDialect>
      <TargetDialect>EN-IN</TargetDialect>
      <TargetDialect>EN-ID</TargetDialect>
      <TargetDialect>EN-IE</TargetDialect>
      <TargetDialect>EN-JM</TargetDialect>
      <TargetDialect>EN-MY</TargetDialect>
      <TargetDialect>EN-NZ</TargetDialect>
      <TargetDialect>EN-PH</TargetDialect>
      <TargetDialect>EN-SG</TargetDialect>
      <TargetDialect>EN-ZA</TargetDialect>
      <TargetDialect>EN-TT</TargetDialect>
      <TargetDialect>EN-GB</TargetDialect>
      <TargetDialect>EN-US</TargetDialect>
      <TargetDialect>EN-ZW</TargetDialect>
    </TargetDialects>
    <!--QueryUri is the address of this dictionary's XML web service (which is used to put definitions in additional contexts, such as the spelling checker.)-->
    <QueryUri DefaultValue="http://christophernlg/ExampleDictionary/WebService.asmx/Define?word="/>
    <!--Citation Text, Dictionary Name, and Dictionary Home Page will be combined to form the citation line (for example, this would produce "Examples by: Microsoft", where "Microsoft" is a hyperlink to http://www.microsoft.com).-->
    <CitationText DefaultValue="Examples by: " />
    <DictionaryName DefaultValue="Microsoft" />
    <DictionaryHomePage DefaultValue="http://www.microsoft.com" />
  </Dictionary>
</OfficeApp>
```

辞書アドインのマニフェスト ファイルの作成に固有の **Dictionary** 要素とその子要素については、以下のセクションで説明します。マニフェスト ファイルのその他の要素の詳細については、「[Office アドイン XML マニフェスト](../../docs/overview/add-in-manifests.md)」を参照してください。


### <a name="dictionary-element"></a>Dictionary 要素


辞書アドインの設定を指定します。

 **親要素**

 `<OfficeApp>`

 **子要素**

 `<TargetDialects>`,  `<QueryUri>`,  `<CitationText>`,  `<DictionaryName>`,  `<DictionaryHomePage>`

 **注釈**

**Dictionary** 要素とその子要素は、辞書アドインを作成するときに作業ウィンドウ アドインのマニフェストに追加されます。


#### <a name="targetdialects-element"></a>TargetDialects 要素


この辞書がサポートする地域言語を指定します。必須 (辞書アドインの場合)。

 **親要素**

 `<Dictionary>`

 **子要素**

 `<TargetDialect>`

 **注釈**

**TargetDialects** 要素とその子要素は、辞書に含める地域言語のセットを指定します。たとえば、スペイン語 (メキシコ) とスペイン語 (ペルー) の両方、ただしスペイン語 (スペイン) は含まないというような指定を、この要素で行うことができます。このマニフェストでは、複数の言語 (たとえば、スペイン語と英語) は指定しないでください。異なる言語は、別個の辞書として発行してください。

 **例**




```XML
<TargetDialects>
  <TargetDialect>EN-AU</TargetDialect>
  <TargetDialect>EN-BZ</TargetDialect>
  <TargetDialect>EN-CA</TargetDialect>
  <TargetDialect>EN-029</TargetDialect>
  <TargetDialect>EN-HK</TargetDialect>
  <TargetDialect>EN-IN</TargetDialect>
  <TargetDialect>EN-ID</TargetDialect>
  <TargetDialect>EN-IE</TargetDialect>
  <TargetDialect>EN-JM</TargetDialect>
  <TargetDialect>EN-MY</TargetDialect>
  <TargetDialect>EN-NZ</TargetDialect>
  <TargetDialect>EN-PH</TargetDialect>
  <TargetDialect>EN-SG</TargetDialect>
  <TargetDialect>EN-ZA</TargetDialect>
  <TargetDialect>EN-TT</TargetDialect>
  <TargetDialect>EN-GB</TargetDialect>
  <TargetDialect>EN-US</TargetDialect>
  <TargetDialect>EN-ZW</TargetDialect>
</TargetDialects>
```


#### <a name="targetdialect-element"></a>TargetDialect 要素


この辞書がサポートする地域言語を指定します。必須 (辞書アドインの場合)。

 **親要素**

 `<TargetDialects>`

 **注釈**

RFC1766 の `language` タグの形式 (たとえば EN-US) で地域言語の値を指定します。

 **例**




```XML
<TargetDialect>EN-US</TargetDialect>
```


#### <a name="queryuri-element"></a>QueryUri 要素


辞書のクエリ サービスのエンドポイントを指定します。辞書アドインでは必須です。

 **親要素**

 `<Dictionary>`

 **注釈**

これは、辞書プロバイダーの XML Web サービスの URI です。この URI の末尾に、適切にエスケープされたクエリが付加されます。 

 **例**




```XML
<QueryUri DefaultValue="http://msranlc-lingo1/proof.aspx?q="/>
```


#### <a name="citationtext-element"></a>CitationText 要素


引用で使用するテキストを指定します。辞書アドインでは必須です。

 **親要素**

 `<Dictionary>`

 **注釈**

この要素では、Web サービスから返されたコンテンツの下の行に表示される引用テキストの冒頭部分を指定します (たとえば "Results by: "、"Powered by: " など)。

この要素では、**Override** 要素を使用して、別のロケールに対応する値を指定できます。たとえば、スペイン語版の SKU の Office を利用しているユーザーが英語の辞書を使用している場合に、引用行を "Results by: Bing" ではなく "Resultados por: Bing" と表示できます。別のロケールに対応する値を指定する方法の詳細については、「[Office アドイン XML マニフェスト](../../docs/overview/add-in-manifests.md)」の「別のロケールに対応する設定値の指定」を参照してください。

 **例**




```XML
<CitationText DefaultValue="Results by: " />
```


#### <a name="dictionaryname-element"></a>DictionaryName 要素


この辞書の名前を指定します。辞書アドインでは必須です。

 **親要素**

 `<Dictionary>`

 **注釈**

この要素では、引用テキスト内のリンク テキストを指定します。引用テキストは、Web サービスから返されたコンテンツの下の行に表示されます。

この要素では、別のロケールに対応する値も指定できます。

 **例**




```XML
<DictionaryName DefaultValue="Bing Dictionary" />
```


#### <a name="dictionaryhomepage-element"></a>DictionaryHomePage 要素


辞書のホーム ページの URL を指定します。辞書アドインでは必須です。

 **親要素**

 `<Dictionary>`

 **注釈**

この要素では、引用テキスト内のリンクの URL を指定します。引用テキストは、Web サービスから返されたコンテンツの下の行に表示されます。

この要素では、別のロケールに対応する値も指定できます。

 **例**




```XML
<DictionaryHomePage DefaultValue="http://www.bing.com" />
```


### <a name="creating-a-dictionary-add-ins-html-user-interface"></a>辞書アドインの HTML ユーザー インターフェイスの作成


次の 2 つの例は、デモの辞書アドインの UI の HTML ファイルと CSS ファイルを示します。アドインの作業ウィンドウでの UI の表示については、コードの下の図 6 を参照してください。Dictionary.js ファイル内の JavaScript の実装でこの HTML の UI のプログラミング ロジックを実現する方法については、次の「JavaScript の実装の記述」を参照してください。


```HTML
<!DOCTYPE html>
<html>

<head>
<meta http-equiv="X-UA-Compatible" content="IE=Edge"/>

<!--The title will not be shown but is supplied to ensure valid HTML.-->
<title>Example Dictionary</title>

<!--Required library includes.-->
<script type="text/javascript" src="http://ajax.microsoft.com/ajax/4.0/1/MicrosoftAjax.js"></script>
<script type="text/javascript" src="office.js"></script>

<!--Optional library includes.-->
<script type="text/javascript" src="http://ajax.aspnetcdn.com/ajax/jQuery/jquery-1.5.1.js"></script>

<!--App-specific CSS and JS.-->
<link rel="Stylesheet" type="text/css" href="style.css" />
<script type="text/ecmascript" src="dictionary.js"></script>
</head>

<body>
<div id="mainContainer">
    <div id="header">
        <span id="headword"></span>
        <span id="pronunciation">(<a id="pronunciationLink">Pronounce</a>)</span>
    </div>
    <ol id="definitions">
    </ol>
    <div id="SeeMore">
    <a id="SeeMoreLink">See More...</a>
    </div>
</div>
</body>

</html>
```

次の例は Style.css の内容を示しています。




```
#mainContainer
{
    font-family: Segoe UI;
    font-size: 11pt;
}

#headword
{
    font-family: Segoe UI Semibold;
    color: #262626;
}

#pronunciation
{
    margin-left: 2px;
    margin-right: 2px;
}

#definitions
{
    font-size: 8.5pt;
}
a
{
    font-size: 8pt;
    color: #336699;
    text-decoration: none;
}
a:visited
{
    color: #993366;
}
a:hover, a:active
{  
    text-decoration: underline;
}
```


**図 6.辞書 UI のデモ**

![辞書 UI のデモ](../../images/DictionaryAgave06.jpg)


### <a name="writing-the-javascript-implementation"></a>JavaScript の実装の記述


次の例は、Dictionary.js ファイル内の JavaScript の実装を示しています。アドインの HTML ページから呼び出されるこのコードによって、デモの辞書アドインのプログラミング ロジックが実現されます。このスクリプトでは、上で説明した XML Web サービスを再利用しています。例の Web サービスと同じディレクトリにスクリプトを配置することによって、そのサービスから定義が取得されますこのスクリプトは、OfficeDefinitions に準拠したパブリック XML Web サービスで使用することもできます。ファイルの冒頭部分にある `xmlServiceURL` 変数を変更し、発音の Bing API キーを、適切に登録されたキーで置き換えます。

この実装で呼び出している JavaScript API for Office (Office.js) の主なメンバーを次に示します。


- **Office** オブジェクトの [initialize](../../reference/shared/office.initialize.md) イベント。これは、アドイン コンテキストの初期化時に発生し、アドインの対象のドキュメントを表す [Document](../../reference/shared/document.md) オブジェクトのインスタンスへのアクセスを提供します。
    
- **Document** オブジェクトの [addHandlerAsync](../../reference/shared/document.addhandlerasync.md) メソッド。これは **initialize** 関数で呼び出されて、ドキュメントの [SelectionChanged](../../reference/shared/document.selectionchanged.event.md) イベントのイベント ハンドラーを追加して、ユーザーの選択範囲の変更をリッスンします。
    
- **Document** オブジェクトの [getSelectedDataAsync](../../reference/shared/document.getselecteddataasync.md) メソッド。これは、**SelectionChanged** イベント ハンドラーの発生時に `tryUpdatingSelectedWord()` 関数で呼び出されて、ユーザーが選択した語句の取得、プレーン テキストへの変換、および非同期コールバック関数 `selectedTextCallback` を実行します。
    
- **getSelectedDataAsync** メソッドの _callback_ 引数で渡した非同期コールバック関数 `selectTextCallback` が実行されると、コールバックが戻った時点で、ユーザーが選択したテキストの値を取得します。この値は、返された **AsyncResult** オブジェクトの [value](../../reference/shared/asyncresult.status.md) プロパティを使用することによって、コールバックの _selectedText_ 引数 (型は [AsyncResult](../../reference/shared/asyncresult.md)) から取得します。
    
- `selectedTextCallback` 関数の残りのコードでは、XML Web サービスへのクエリで定義を取得します。また、Microsoft Translator API を呼び出して、選択した語句の発音が入った .wav ファイルの URL も取得します。
    
- Dictionary.js の残りのコードでは、アドインの HTML の UI に定義のリストと発音のリンクを表示します。
    



```
// The document the dictionary add-in is interacting with.
var _doc; 
// The last looked-up word, which is also the currently displayed word.
var lastLookup; 
// For demo purposes only!! Get an AppID if you intend to use the Pronunciation service for your feature.
var appID="3D8D4E1888B88B975484F0CA25CDD24AAC457ED8"; 

// The base URL for the OfficeDefinitions-conforming XML web service to query for definitions.
var xmlServiceUrl = "WebService.asmx/Define?Word="; 

// Initialize the add-in. 
// The initialize function is required for all add-ins.
Office.initialize = function (reason) {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    // Store a reference to the current document.
    _doc = Office.context.document; 
    // Check whether text is already selected.
    tryUpdatingSelectedWord(); 
    _doc.addHandlerAsync("documentSelectionChanged", tryUpdatingSelectedWord); //Add a handler to refresh when the user changes selection.
    });
}

// Executes when event is raised on user's selection changes, and at initialization time. 
// Gets the current selection and passes that to asynchronous callback method.
function tryUpdatingSelectedWord() {
    _doc.getSelectedDataAsync(Office.CoercionType.Text, selectedTextCallback); 
}

// Async callback that executes when the add-in gets the user's selection.
// Determines whether anything should be done. If so, it makes requests that will be passed to various functions.
function selectedTextCallback(selectedText) {
    selectedText = $.trim(selectedText.value);
    // Be sure user has selected text. The SelectionChanged event is raised every time the user moves the cursor, even if no selection.
    if (selectedText != "") { 
        // Check whether user selected the same word the pane is currently displaying to avoid unnecessary web calls.
        if (selectedText != lastLookup) { 
            // Update the lastLookup variable.
            lastLookup = selectedText; 
            // Set the "headword" span to the word you looked up.
            $("#headword").text(selectedText); 
            // AJAX request to get definitions for the selected word; pass that to refreshDefinitions.
            $.ajax(xmlServiceUrl + selectedText, { dataType: 'xml', success: refreshDefinitions, error: errorHandler }); 
            // AJAX request to the Microsoft Translator APIs. Gets the URL of a WAV file with pronunciation, which is passed to refreshPronunciation. See http://www.microsofttranslator.com/dev for details.
            $.ajax("http://api.microsofttranslator.com/V2/Ajax.svc/Speak?oncomplete=refreshPronunciation&amp;appId=" + appID + "&amp;text=" + selectedText + "&amp;language=en-us", { dataType: 'script', success: null, error: errorHandler }); 
        }
    }
}

// This function is called when the add-in gets back the definitions target word.
// It removes the old definitions and replaces them with the definitions for the current word.
// It also sets the "See More" link.
function refreshDefinitions(data, textStatus, jqXHR) {
    $(".definition").remove();
    // Make a new list item for each returned definition that was returned, set the CSS class, and append it to the definitions div.
    $(data).find("Definition").each(function () {
        $(document.createElement("li")).text($(this).text()).addClass("definition").appendTo($("#definitions"));
    });
    $("#SeeMoreLink").attr("href", $(data).find("SeeMoreURL").text()); //Change the "See More" link to direct to the correct URL.
}

// This function is called when the add-in gets back the link to the pronunciation
// to set the "Pronounce" link to the URL of the .WAV file.
function refreshPronunciation(data) {
    $("#pronunciationLink").attr("href", data);
}

// Basic error handler that writes to a div with id='message'.
function errorHandler(jqXHR, textStatus, errorThrown) {
    document.getElementById('message').innerText += errorThrown;
}

```

