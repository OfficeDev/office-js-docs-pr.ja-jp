---
title: Word アドインで Office Open XML を使用するタイミングと方法を理解する
description: Word アドインで Office Open XML をいつどのように使用するかの概要。
ms.date: 07/15/2022
ms.localizationpriority: medium
ms.openlocfilehash: b3a20c18677b6af11de835a996fc8eaf912a7770
ms.sourcegitcommit: df7964b6509ee6a807d754fbe895d160bc52c2d3
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/20/2022
ms.locfileid: "66889549"
---
# <a name="understand-when-and-how-to-use-office-open-xml-in-your-word-add-in"></a>Word アドインで Office Open XML を使用するタイミングと方法を理解する

**提供元:** Stephanie Krieger, Microsoft Corporation | Juan Balmori Labra, Microsoft Corporation

Word で実行する Office アドインをビルドする場合、Office JavaScript API (Office.js) には、ドキュメント コンテンツの読み取りと書き込みのためのいくつかの形式が既に用意されていることがわかります。 そのような書式は強制タイプと呼ばれ、プレーン テキスト、表、HTML、Office Open XML が含まれています。

それでは、イメージ、書式設定された表、グラフ、あるいは書式設定だけされたテキストをドキュメントに追加する必要があるとき、何を選択しますか。 [WordApi 要件セット](/javascript/api/requirement-sets/word/word-api-requirement-sets)から使用可能な API から始めて、必要なものを提供するかどうかを確認します。 例については、Word の[スクリプト ラボ](https://appsource.microsoft.com/product/office/wa104380862)で[書式設定されたテキスト](https://github.com/OfficeDev/office-js-snippets/blob/prod/samples/word/25-paragraph/insert-formatted-text.yaml) コード スニペットを挿入するを参照してください。

API がまだ使用できない場合は、HTML を使用して、画像などのリッチ コンテンツの種類を挿入できます。 シナリオによっては、コンテンツで使用できる書式設定および配置に関する制限など、HTML 強制型変換にとって不利な点が存在する場合があります。

Office Open XML は、Word 文書 (.docx、.dotx など) が記述される言語であるため、Word 文書にユーザーが追加できるコンテンツであれば、ユーザーにとって適用可能なほぼすべてのタイプの書式設定で、事実上あらゆるタイプのコンテンツを挿入できます。 処理する Office Open XML マークアップを判別するのは、想像しているよりも簡単です。

> [!NOTE]
> Office Open XML は PowerPoint と Excel (ならびに、Office 2013 以降、Visio) ドキュメントの言語です。 ただし、現在のところは、Word 用に作成された Office アドインでのみ、Office Open XML としてコンテンツを型強制できます。 完全な言語リファレンス ドキュメントを含む Office Open XML の詳細については、 [こちらも参照](#see-also) してください。

まず、Office Open XML 強制を使用して挿入できるコンテンツ タイプの一部を見てみましょう。
次の例のいずれかを Word に挿入するために必要な Office Open XML マークアップとOffice.jsコードを含む、 [Word アドインで Open XML を読み込んで記述](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/word-add-in-load-and-write-open-xml)するコード サンプルをダウンロードします。

> [!NOTE]
> この記事では、 **コンテンツ タイプ** と **リッチ コンテンツ** という用語は、Word 文書に挿入できるリッチ コンテンツの種類を指します。

*図 1. 書式が直接設定されたテキスト*

![直接書式設定が適用されたテキスト。](../images/office15-app-create-wd-app-using-ooxml-fig01.png)

直接書式を使用して、ユーザーのドキュメント内の既存の書式設定に関係なく、テキストの外観を正確に指定します。

*図 2. スタイルを使用して書式設定されたテキスト*

![段落スタイルで書式設定されたテキスト。](../images/office15-app-create-wd-app-using-ooxml-fig02.png)

スタイルを使用して、挿入したテキストの外観をユーザーのドキュメントと自動的に調整します。

*図 3. 単純なイメージ*

![サンプル ロゴ。](../images/office15-app-create-wd-app-using-ooxml-fig03.png)

Office でサポートされているイメージ形式を挿入する場合と同じ方法を使用します。

*図 4. ピクチャのスタイルと効果を使用して書式設定されたイメージ*

![Word 内の書式設定された画像。](../images/office15-app-create-wd-app-using-ooxml-fig04.png)

高品質の書式設定と効果をイメージに追加するとき、必要なマークアップは予想よりずっと少なくなります。

*図 5. コンテンツ コントロール*

![境界線コンテンツ コントロール内のテキスト。](../images/office15-app-create-wd-app-using-ooxml-fig05.png)

アドインでコンテンツ コントロールを使用して、選択範囲ではなく、指定された (バインドされた) 場所にコンテンツを追加します。

*図 6. ワードアートが書式設定されたテキスト ボックス*

![ワードアート テキスト効果で書式設定されたテキスト。](../images/office15-app-create-wd-app-using-ooxml-fig06.png)

Word では、テキスト ボックス (図参照) 内のテキストに、または通常の本文テキストにテキスト効果を利用できます。

*図 7. 図形*

![Word の描画図形。](../images/office15-app-create-wd-app-using-ooxml-fig07.png)

テキストと書式設定効果の有無にかかわらず、組み込みまたはカスタムの描画図形を挿入します。

*図 8. 直接書式設定の表*

![Word 内の書式設定されたテーブル。](../images/office15-app-create-wd-app-using-ooxml-fig08.png)

テキストの書式設定、罫線、網かけ、セルのサイズ設定、または必要なテーブルの書式設定を含めます。

*図 9. 表スタイルを使用して書式設定されたテーブル*

![Word の表スタイルで書式設定されたテーブル。](../images/office15-app-create-wd-app-using-ooxml-fig09.png)

テキストに段落スタイルを使用するのと同じくらい簡単に、組み込みまたはカスタムの表スタイルを使用します。

*図 10. SmartArt 図*

![Word 内の動的な SmartArt 図。](../images/office15-app-create-wd-app-using-ooxml-fig10.png)

Office には、さまざまな SmartArt ダイアグラム レイアウトが用意されています (また、Office Open XML を使用して独自のレイアウトを作成することもできます)。

*図 11. グラフ*

![Word 内のグラフ。](../images/office15-app-create-wd-app-using-ooxml-fig11.png)

Word 文書に Excel グラフをライブ グラフとして挿入できます。つまり、Word 用アドインでも使用できます。
前の例でわかるように、Office Open XML 強制を使用して、ユーザーが独自のドキュメントに挿入できるあらゆる種類のコンテンツを本質的に挿入できます。
必要な Office Open XML マークアップを取得するには、2 つの簡単な方法があります。 リッチ コンテンツをそれ以外の空の Word 文書に追加してから、ファイルを Word XML ドキュメント形式で保存するか、 [getSelectedDataAsync](/javascript/api/office/office.document#office-office-document-getselecteddataasync-member(1)) メソッドでテスト アドインを使用してマークアップを取得します。 いずれの方法でも基本的には同じ結果が得られます。

> [!NOTE]
> Office Open XML ドキュメントは、実際にはドキュメントの内容を表すファイルの圧縮パッケージです。 Word XML ドキュメント形式でファイルを保存すると、Office Open XML パッケージ全体が 1 つの XML ファイルにフラット化されます。これは、Office Open XML マークアップの取得に使用 `getSelectedDataAsync` する場合にも使用できます。

Word から XML 形式にファイルを保存する場合は、[名前を付けて保存] ダイアログ ボックスの [名前を付けて保存] ボックスの一覧に、.xml書式ファイルの 2 つのオプションがあることに注意してください。 Word 2003 オプションではなく、必ず [**Word XML ドキュメント**] を選択してください。
[Word-Add-in-Get-Set-EditOpen-XML](https://github.com/OfficeDev/Word-Add-in-Get-Set-EditOpen-XML) という名前のコード サンプルをダウンロードします。これは、マークアップを取得してテストするためのツールとして使用できます。
これですべてでしょうか。 そうではありません。 多くのシナリオについては、先ほどのメソッドのいずれかを使用して表示される、完全な、圧縮 Office Open XML 結果を使用できるし、それは機能するでしょう。 良いニュースは、おそらくそのマークアップのほとんどを必要としないということです。
Office Open XML マークアップを初めて見る多くのアドイン開発者の 1 人である場合は、最も単純なコンテンツに対して取得する大量のマークアップを理解しようとしている場合は、圧倒されているように見えますが、そうである必要はありません。
このトピックでは、Office アドイン開発者コミュニティから聞いたいくつかの一般的なシナリオを使用して、アドインで使用するために Office Open XML を簡略化する手法を紹介します。 前に示した一部の種類のコンテンツのマークアップと、Office Open XML ペイロードを最小限に抑えるために必要な情報について説明します。 また、作業中の選択範囲でドキュメントにリッチ コンテンツを挿入するために必要なコードと、Office Open XML をバインド オブジェクトと共に使用して、指定した場所でコンテンツを追加または置き換える方法についても説明します。

## <a name="explore-the-office-open-xml-document-package"></a>Office Open XML ドキュメント パッケージを調べる

[getSelectedDataAsync](/javascript/api/office/office.document#office-office-document-getselecteddataasync-member(1)) を使用して、選択したコンテンツの Office Open XML を取得する場合 (または Word XML ドキュメント形式でドキュメントを保存する場合) は、選択したコンテンツを記述するマークアップだけではありません。これは、ほぼ確実に必要のない多くのオプションと設定を備えたドキュメント全体です。 事実、作業ウィンドウ アドインを含むドキュメントからそのメソッドを使用すると、取得するマークアップにはその作業ウィンドウ アプリまでもが含まれています。

単純な Word ドキュメント パッケージであっても、実際のコンテンツだけではなく、ドキュメント プロパティ、スタイル、テーマ (書式設定)、Web 設定、フォントなどが含まれています。

たとえば、先ほど図 1 で表示されていたように、直接書式が設定されたテキストの 1 段落だけを挿入するとします。 書式設定されたテキスト用に Office Open XML を取得  `getSelectedDataAsync`すると、大量のマークアップが表示されます。 このマークアップには、図 13 に示すように、ドキュメント全体を表すパッケージ要素が含まれています。これには、いくつかの部分 (一般的にドキュメント パーツと呼ばれるか、Office Open XML ではパッケージ パーツと呼ばれます) が含まれています。 各パーツは、パッケージ内の個別のファイルを表します。

> [!TIP]
> メモ帳などのテキスト エディターで Office Open XML マークアップを編集します。 Visual Studio で開く場合は、 **編集>高度な>文書の書式設定** (Ctrl + K、Ctrl + D) を使用して、簡単に編集できるようにパッケージを書式設定します。 次に図 12 で示すように、ドキュメントの一部またはセクションを折りたたんだり、展開したりして、Office Open XML パッケージのコンテンツをより簡単に確認したり、編集したりすることができます。 各ドキュメントの部分は **pkg:part** タグから始まります。

*図 12. Visual Studio で簡単に編集するためのパッケージ パーツの圧縮と展開*

![Visual Studio でパッケージ パーツの Office Open XML コード スニペット。](../images/office15-app-create-wd-app-using-ooxml-fig12.png)

*図 13. 基本的な Word Office Open XML ドキュメント パッケージに含まれるパーツ*

![パッケージ部分の Office Open XML コード スニペット。](../images/office15-app-create-wd-app-using-ooxml-fig13.png)

これだけたくさんマークアップが存在するのに、書式設定されたテキスト例を挿入するために実際に必要な要素は .rels パーツと document.xml パーツだけであることを知ったら驚かれることでしょう。

> [!NOTE]
> パッケージ タグの上にあるマークアップの 2 行 (バージョンと Office プログラム ID の XML 宣言) は、Office Open XML 強制型変換を使用するときに採用されます。そのため、自分で追加する必要はありません。編集したマークアップを Word ドキュメントとして開き、テストする場合、そのまま利用してください。

このトピックの冒頭に示す他のいくつかの種類のコンテンツでは、(図 13 を超えて) 追加の部分も必要であり、このトピックの後半で取り上げます。 一方、図 13 に示すほとんどの部分が Word ドキュメント パッケージのマークアップに表示されるため、これらの各パーツの用途と必要なタイミングを簡単に説明します。

- パッケージ タグ内部の最初のパーツは .rels ファイルです。これはパッケージの上位レベルのパーツ (一般的には、ドキュメント プロパティ、存在する場合はサムネイル、ドキュメント本文) 間のリレーションシップを定義します。ドキュメント本文 (コンテンツが置かれる場所) とドキュメント パッケージとのリレーションシップを定義する必要があるので、このパーツのコンテンツの一部は常にマークアップで必要です。

- document.xml (本文) パーツに必要な追加パーツが存在する場合、そのリレーションシップが document.xml.rels パーツにより定義されます。

   > [!IMPORTANT]
   > パッケージの .rels ファイル (top-level .rels、document.xml.rels、および特定のタイプのコンテンツで遭遇するその他のファイル) はガイドとして使用できる非常に重要なツールであり、Office Open XML パッケージをすばやく編集できます。編集方法に関する詳細は、本トピックの後半にある「[独自のマークアップを作成する: ベスト プラクティス](#create-your-own-markup-best-practices)」を参照してください。

- document.xml パーツはドキュメントの本文のコンテンツです。コンテンツが表示される部分であるため、もちろんこのパーツの要素は必要です。しかしながら、このパーツに表示されるすべてが必要なわけではありません。これについては後で詳しく説明します。

- 多くのパーツは、Office Open XML 強制型変換を利用してドキュメントにコンテンツを挿入するときに Set メソッドにより自動的に無視されます。そのため、削除してもかまいません。このようなパーツには、theme1.xml ファイル (ドキュメントの書式設定テーマ)、ドキュメントのプロパティ パーツ (コア、アドイン、サムネイル)、設定ファイル (settings、webSettings、fontTable を含む) が含まれます。

- 図 1 の例では、テキストの書式設定が直接適用されています (つまり、各フォントおよび段落の書式設定が個別に適用されます)。 ただし、先の図 2 のように、スタイルを使用する場合 (たとえば、変換先ドキュメントの見出し 1 スタイルの書式設定をテキストに自動的に適用する場合)、styles.xml パーツとそのリレーションシップ定義が必要になります。 詳細については、「 [追加の Office Open XML パーツを使用するオブジェクトを追加する」](#add-objects-that-use-additional-office-open-xml-parts)セクションを参照してください。

## <a name="insert-document-content-at-the-selection"></a>選択範囲にドキュメント コンテンツを挿入する

それでは、図 1 の書式設定テキスト例に必要な最小 Office Open XML マークアップと、ドキュメントでアクティブになっている選択箇所に挿入するために必要な JavaScript について見てみましょう。

### <a name="simplified-office-open-xml-markup"></a>単純化された Office Open XML マークアップ

前のセクションで説明したように、次に示す Office Open XML の例を編集して、必要なドキュメント パーツのみを残し、それらの各部分内に必要な要素のみを残しました。 このトピックの次のセクションでは、マークアップを自分で編集する方法について説明します (また、ここで残っている部分についてもう少し説明します)。

```XML
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml" pkg:padding="512">
    <pkg:xmlData>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
      </Relationships>
    </pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" >
        <w:body>
          <w:p>
            <w:pPr>
              <w:spacing w:before="360" w:after="0" w:line="480" w:lineRule="auto"/>
              <w:rPr>
                <w:color w:val="70AD47" w:themeColor="accent6"/>
                <w:sz w:val="28"/>
              </w:rPr>
            </w:pPr>
            <w:r>
              <w:rPr>
                <w:color w:val="70AD47" w:themeColor="accent6"/>
                <w:sz w:val="28"/>
              </w:rPr>
              <w:t>This text has formatting directly applied to achieve its font size, color, line spacing, and paragraph spacing.</w:t>
            </w:r>
          </w:p>
        </w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>
```

> [!NOTE]
> ここに示すマークアップを XML ファイルに追加し、バージョンと mso アプリケーションの XML 宣言タグ (図 13) をファイルの先頭に追加すると、そのファイルを Word ドキュメントとして Word で開くことができます。 または、これらのタグを使用せずに、Word で **ファイル** > を **開いて開** くこともできます。 Word のタイトル バーに **互換モード** が表示されます。これは Word 文書であることを示す設定を削除したためです。 このマークアップを既存の Word 文書に追加しているため、コンテンツにはまったく影響しません。

### <a name="javascript-for-using-setselecteddataasync"></a>setSelectedDataAsync を使用するための JavaScript

前の Office Open XML をソリューションからアクセスできる XML ファイルとして保存したら、次の関数を使用して、Office Open XML 強制を使用してドキュメント内の書式設定されたテキスト コンテンツを設定します。

この関数では、最後の行を除くすべての行が、関数の最後の [setSelectedDataAsync](/javascript/api/office/office.document#office-office-document-setselecteddataasync-member(1)) メソッド呼び出しで使用するために保存されたマークアップを取得するために使用されていることに注意してください。 `setSelectedDataASync` は、挿入するコンテンツと強制型を指定することだけが必要です。

> [!NOTE]
> *yourXMLfilename* をソリューションに保存した XML ファイルの名前とパスで置き換えます。 ソリューションに XML ファイルを含める場所やコード内で XML ファイルを参照する方法がわからない場合は、その例と、ここに示すマークアップと JavaScript の実際の例については、Word アドインのコード サンプル [で Open XML を読み込んで記述](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/word-add-in-load-and-write-open-xml) する方法を参照してください。

```js
function writeContent() {
    const myOOXMLRequest = new XMLHttpRequest();
    let myXML;
    myOOXMLRequest.open('GET', 'yourXMLfilename', false);
    myOOXMLRequest.send();
    if (myOOXMLRequest.status === 200) {
        myXML = myOOXMLRequest.responseText;
    }
    Office.context.document.setSelectedDataAsync(myXML, { coercionType: 'ooxml' });
}
```

## <a name="create-your-own-markup-best-practices"></a>独自のマークアップを作成する: ベスト プラクティス

それでは、先ほどの書式設定テキスト例を挿入するために必要なマークアップについて詳しく見てみましょう。

この例では、.rels と document.xml 以外のすべてのドキュメント パーツをパッケージから削除することから始めます。 次に、これら 2 つの必要な部分を編集して、さらに簡略化します。

> [!IMPORTANT]
> .rels パーツは、パッケージに含めるパーツをすばやく判断し、完全に削除できるパーツ (つまり、コンテンツに関係ない、またはコンテンツに参照されないパーツ) を判断するためのマップとして使用します。すべてのドキュメント パーツについて、パッケージにリレーションシップを定義する必要があることと、それらのリレーションシップが .rels ファイルに表示されることに注意してください。そのため、.rels、document.xml.rels、コンテンツ固有 .rels ファイルのいずれかにそれらのすべてが一覧表示される必要があります。

次のマークアップは、編集前に必要な .rels パーツを示しています。 アドインとコア ドキュメントのプロパティ パーツとサムネイル パーツを削除するため、それらのリレーションシップも .rels から削除する必要があります。 document.xml のリレーションシップのみが残ることに注意してください (次の例で関係 ID が "rID1" のもの)。

```XML
<pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml" pkg:padding="512">
  <pkg:xmlData>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
      <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail" Target="docProps/thumbnail.emf"/>
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
      <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
    </Relationships>
  </pkg:xmlData>
</pkg:part>
```

> [!IMPORTANT]
> パッケージから完全に削除するパーツのリレーションシップ (つまり、**Relationship** タグ) を削除します。対応するリレーションシップなしでパーツを追加するか、パーツを除外してもパッケージにそのリレーションシップを残した場合、エラーが発生します。

次のマークアップは、書式設定されたテキスト コンテンツのサンプルを含む、編集前の document.xml パーツを示しています。

```XML
<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document mc:Ignorable="w14 w15 wp14" xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
        <w:body>
          <w:p>
            <w:pPr>
              <w:spacing w:before="360" w:after="0" w:line="480" w:lineRule="auto"/>
              <w:rPr>
                <w:color w:val="70AD47" w:themeColor="accent6"/>
                <w:sz w:val="28"/>
              </w:rPr>
            </w:pPr>
            <w:r>
              <w:rPr>
                <w:color w:val="70AD47" w:themeColor="accent6"/>
                <w:sz w:val="28"/>
              </w:rPr>
              <w:t>This text has formatting directly applied to achieve its font size, color, line spacing, and paragraph spacing.</w:t>
            </w:r>
            <w:bookmarkStart w:id="0" w:name="_GoBack"/>
            <w:bookmarkEnd w:id="0"/>
          </w:p>
          <w:p/>
          <w:sectPr>
            <w:pgSz w:w="12240" w:h="15840"/>
            <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
            <w:cols w:space="720"/>
          </w:sectPr>
        </w:body>
      </w:document>
    </pkg:xmlData>
</pkg:part>
```

document.xmlはコンテンツを配置する主要なドキュメント パーツであるため、その部分を簡単に確認してください。 (図 14 は、この一覧に従って、ここで説明するコア コンテンツと書式設定タグの一部が Word 文書に表示される内容とどのように関連しているかを示す視覚的な参照を示しています)。

- 冒頭の **w:document** タグには、いくつかの名前空間 (**xmlns**) のリストが含まれます。これらの名前空間の多くは特定のタイプのコンテンツを参照するため、お使いのコンテンツに関連する場合にのみ必要です。

    ドキュメント パーツ全体のタグのプレフィックスが名前空間を参照していることに注意してください。 この例では、document.xml パーツ全体のタグで使用されるプレフィックスは **w:** のみであるため、開 **いている w:document** タグに残す必要がある名前空間は **xmlns:w** のみです。

    > [!TIP]
    > マークアップを Visual Studio で編集する場合、何らかのパーツの名前空間を削除したら、そのパーツのすべてのタグを確認します。マークアップに必要な名前空間を削除した場合、影響を受けたタグの関連するプレフィックスに赤い波下線が表示されます。 **xmlns:mc** 名前空間を削除する場合、名前空間の一覧に先行する **mc:Ignorable** 属性も削除する必要があります。

- 冒頭の body タグの内部に、段落タグ (**w:p**) が表示され、この例のサンプル コンテンツが表示されます。

- **w:pPr** タグには、段落の前後の空間、段落の位置合わせ、字下げなど、直接適用される段落書式設定のプロパティが含まれます (直接書式設定は、スタイルの一部としてではなく、コンテンツに個別に適用する属性を参照します)。またこのタグには、段落全体に適用されるフォントの直接書式設定が、サンプルに設定されたフォントの色とサイズを含む、入れ子になった **w:rPr** (run プロパティ) タグ内に含まれます。

   > [!NOTE]
   > Word Office Open XML マークアップ内のフォント サイズおよび他のいくつかの書式設定では、実際の 2 倍のサイズのように表示されることがあります。これは、段落や行間、および上記のマークアップに表示される書式プロパティが twip (ポイントの 20 分の 1) 単位で指定されているためです。Office Open XML で使用するコンテンツの種類に応じて、英語のメートル法 (1 インチあたり 914,400 Emu) を含むいくつかの追加単位が使用されていることがあります。これは、Office Art (drawingML) の一部の値に使用されます。また実際の 100,000 倍の値が使用されていることがあります。この値は、drawingML と PowerPoint のマークアップの両方で使用されます。PowerPoint でも一部の値は実際の 100 倍の値として表されますが、Excel では通常実際の値が使用されます。

- 段落内で、サンプル テキストのように、同様のプロパティを持つコンテンツが実行に含まれます (**w:r**)。書式設定またはコンテンツ タイプに変更があるたびに、新しい実行が開始します。(つまり、サンプル テキストの 1 文字だけが太字であれば、それは独自の実行に分類されます。) この例では、コンテンツに 1 つだけのテキスト実行が含まれます。

    このサンプルに含まれる書式設定はフォントの書式設定 (つまり、わずか 1 文字に適用される書式設定) であるため、個別実行のプロパティにも表示されることに注意してください。

- 既定で Word ドキュメントに現れる非表示の "_GoBack" ブックマーク (**w:bookmarkStart** と **w:bookmarkEnd**) のタグにも注意が必要です。マークアップからいつでも、GoBack ブックマークの開始タグと終了タグを削除することができます。

- ドキュメント本体の最後の部分は **w:sectPr** タグまたはセクション プロパティです。このタグには余白やページ方向などの設定が含まれます。**setSelectedDataAsync** を利用して挿入したコンテンツは、既定でターゲット ドキュメントのアクティブなセクションのプロパティを獲得します。そのため、コンテンツにセクションの改行が含まれる (その場合、複数の **w:sectPr** タグが表示されます) のでなければ、このタグを削除できます。

*図 14. document.xml の共通タグと Word ドキュメントのコンテンツならびにレイアウトの関係*

![Word 文書内の Office Open XML 要素。](../images/office15-app-create-wd-app-using-ooxml-fig14.png)

> [!TIP]
> 作成するマークアップでは、w **:rsid** 文字を含むいくつかのタグに別の属性が表示される場合があります。このトピックで使用されている例では表示されません。 These are revision identifiers. They're used in Word for the Combine Documents feature and they're on by default. You'll never need them in markup you're inserting with your add-in and turning them off makes for much cleaner markup. 既存の RSID タグを簡単に削除するか、(次の手順で説明するように) 機能を無効にして、新しいコンテンツのマークアップに追加されないようにすることができます。

Word の共同作成機能 (他の人と同時にドキュメントを編集する機能など) を使用する場合、アドインのマークアップ生成を完了したときに、機能を再度有効にする必要があることに注意してください。

将来作成するドキュメントのために Word の RSID 属性をオフにするには、次のように操作します。

1. Word で、[**ファイル**] を選択し、[**オプション**] を選択します。
1. [Word オプション] ダイアログ ボックスで、[**セキュリティ センター**] を選択し、[**セキュリティ センターの設定**] を選択します。
1. [セキュリティ センター] ダイアログ ボックスで、[**プライバシー オプション**] を選択し、[**組み込みの精度を向上させるためランダムな番号を保存する**] という設定を無効にします。

既存のドキュメントから RSID タグを削除するには、Office Open XML で開いたドキュメントで次のショートカットを試します。

1. ドキュメントの本文にカーソルを置き、**Ctrl キー + Home キー** を押してドキュメントの先頭に移動します。
1. キーボードで、**Space キー**、**Del キー**、**Space キー** を押します。次にドキュメントを保存します。

このパッケージからマークアップの大半を削除した後に、前のセクションで見たような、サンプルのために挿入する必要がある最小限のマークアップが残ります。

## <a name="use-the-same-office-open-xml-structure-for-different-content-types"></a>異なるコンテンツ タイプに同じ Office Open XML 構造を使用する

一部のタイプのリッチ コンテンツでは、前の例で見たような、コンテンツ コントロール、Office 描画図形、テキスト ボックス、表 (スタイルが表に適用されない限り) など、.rels および document.xml コンポーネントのみが必要です。実際は、編集済みの同じパッケージ パーツを再利用したり、document.xml の **body** コンテンツを自分のコンテンツのマークアップと交換したりします。

図 5 から図 8 に示した各コンテンツ タイプの例については、Office Open XML マークアップを確認するには、概要セクションで参照されている [Word アドイン コード サンプルで Open XML の読み込みと書き込みを](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/word-add-in-load-and-write-open-xml) 行います。

次に進む前に、いくつかのコンテンツ タイプと必要な部分を入れ替える方法について、相違点を確認してください。

### <a name="understand-drawingml-markup-office-graphics-in-word-what-are-fallbacks"></a>Word の drawingML マークアップ (Office グラフィックス) を理解する: フォールバックとは何ですか?

図形またはテキスト ボックスのマークアップが予想よりはるかに複雑に見える場合は、その理由があります。 Office 2007 のリリースでは、Office Open XML 形式の導入と、PowerPoint と Excel が完全に採用した新しい Office グラフィックス エンジンの導入が見えました。 2007 リリースでは、Word はそのグラフィックス エンジンの一部のみを組み込み、更新された Excel グラフ エンジン、SmartArt グラフィックス、高度な画像ツールを採用しました。 図形とテキスト ボックスの場合、Word 2007 では引き続き従来の描画オブジェクト (VML) が使用されています。 2010 年のリリースでは、更新された図形と描画ツールを組み込むために、Word がグラフィックス エンジンで追加の手順を実行しました。

そのため、Office Open XML 書式の Word ドキュメントを Word 2007 で開いたときにその図形とテキスト ボックスをサポートするために、図形 (テキスト ボックスを含む) にはフォールバック VML マークアップが必要です。

通常、Word アドインのコード サンプルの [Open XML の読み込みと書き込みに](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/word-add-in-load-and-write-open-xml) 含まれる図形とテキスト ボックスの例については、フォールバック マークアップを削除できます。 文書の保存時に、不足しているフォールバック マークアップが図形に自動的に追加されます。 ただし、すべてのユーザー シナリオが確実にサポートされるようにフォールバック マークアップをそのままにしておいても、まったく問題はありません。

コンテンツに含まれる描画オブジェクトをグループ化している場合、追加の (明らかに繰り返しである) マークアップが表示されますが、これはそのままにしておく必要があります。オブジェクトがグループに追加されると、描画図形のマークアップの部分が重複します。

> [!IMPORTANT]
> テキスト ボックスと描画図形を操作するときは、名前空間を慎重に確認してから document.xml から削除してください (あるいは、別のオブジェクト タイプからマークアップを再利用する場合、document.xml から以前に削除した名前空間が必要であれば、それを戻します)。既定で document.xml に含まれる名前空間の重要な部分は、図画オブジェクトの要件のためにあります。

#### <a name="about-graphic-positioning"></a>グラフィックの位置に関する注記

コード サンプルでは、 [Word アドインと Word アドインの Open XML と](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/word-add-in-load-and-write-open-xml) [Word-Add-In-Get-Set-EditOpen-XML](https://github.com/OfficeDev/Word-Add-in-Get-Set-EditOpen-XML) を読み込んで書き込みます。テキスト ボックスと図形は、さまざまな種類のテキスト の折り返しと配置の設定を使用して設定されます。 (また、これらのコード サンプルの画像の例は、テキストのベースライン上にグラフィック オブジェクトを配置するテキストの書式設定と一致して設定されていることにも注意してください)。

これらのコード サンプルの図形は右と下のページ余白との関連で位置付けられます。相対的位置付けにより、簡単にユーザーの不明なドキュメント設定を調整でき (ユーザーの余白に合わせて調整されるため)、紙のサイズ、方向、余白の設定によってぎこちなく見えることもなくなります。グラフィック オブジェクトを挿入するときに相対的位置付け設定を維持するには、位置付け (Word のアンカー) が保存される段落マーク (w:p) を維持する必要があります。独自の段落マークを追加するのではなく、既存の段落マークにコンテンツを挿入する場合、同じ初期ビジュアルを維持できても、ユーザーのレイアウトに合わせて自動的に調整する位置付けを有効にする多くのタイプの相対参照が失われる可能性があります。

### <a name="work-with-content-controls"></a>コンテンツ コントロールを操作する

コンテンツ コントロールは Word の重要な機能であり、さまざまな方法で Word 用アドインの機能を大幅に強化します。たとえば、選択範囲だけではなく、ドキュメントの指定の場所にコンテンツを挿入できます。

Word では、図 15 のように、コンテンツ コントロールはリボンの [開発] タブにあります。

*図 15. Word の [開発] タブの [コントロール] グループ*

![Word のリボン上のコンテンツ コントロール グループ。](../images/office15-app-create-wd-app-using-ooxml-fig15.png)

Word のコンテンツ コントロールの種類には、リッチ テキスト、プレーン テキスト、ピクチャ、構成要素ギャラリー、チェック ボックス、ドロップダウン リスト、コンボ ボックス、日付の選択、繰り返しセクションがあります。

- 図 15 にある [**プロパティ**] コマンドを使用し、コントロールのタイトルを編集したり、コントロール コンテナーを非表示にするなど、設定を追加します。

- [**デザイン モード**] を有効化し、コントロールのプレースホルダー コンテンツを編集します。

アドインが Word テンプレートで動作する場合は、そのテンプレートにコントロールを含め、コンテンツの動作を強化できます。 また、Word 文書の XML データ バインディングを使用して、コンテンツ コントロールをドキュメント プロパティなどのデータにバインドして、フォーム入力や同様のタスクを簡単に行うこともできます。 ([ **挿入** ] タブの [ **クイック パーツ**] の下にある Word の組み込みのドキュメント プロパティに既にバインドされているコントロールを検索します)。

アドインでコンテンツ コントロールを使用すると、異なるタイプのバインドを使用してアドインで可能なオプションを大幅に拡張できます。アドインの中からコンテンツ コントロールをバインドし、アクティブな選択範囲ではなく、バインドにコンテンツを記述できます。

> [!NOTE]
> Word の XML データ バインドと、アドインを介してコントロールにバインドする機能を混同しないでください。これらはまったく別の機能です。ただし、OOXML 強制型変換を利用してアドイン経由で挿入したコンテンツに指定のコンテンツ コントロールを追加し、アドインのコードを使用してこれらのコントロールにバインドすることができます。

また、XML データ バインディングとOffice.jsの両方がアプリ内のカスタム XML パーツと対話できるため、これらの強力なツールを統合できることにも注意してください。 Office JavaScript API のカスタム XML パーツの使用については、このトピックの「[その他の技術情報](#see-also)」セクションを参照してください。

Word アドインのバインドの使用については、トピックの次のセクションで取り上げます。 まず、アドインを使用してバインドできるリッチ テキスト コンテンツ コントロールを挿入するために必要な Office Open XML の例を見てみましょう。

> [!IMPORTANT]
> リッチ テキスト コントロールは、アドイン内からコンテンツ コントロールにバインドできる唯一のタイプのコンテンツ コントロールです。

```XML
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml" pkg:padding="512">
    <pkg:xmlData>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
      </Relationships>
    </pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" >
        <w:body>
          <w:p/>
          <w:sdt>
              <w:sdtPr>
                <w:alias w:val="MyContentControlTitle"/>
                <w:id w:val="1382295294"/>
                <w15:appearance w15:val="hidden"/>
                <w:showingPlcHdr/>
              </w:sdtPr>
              <w:sdtContent>
                <w:p>
                  <w:r>
                  <w:t>[This text is inside a content control that has its container hidden. You can bind to a content control to add or interact with content at a specified location in the document.]</w:t>
                </w:r>
                </w:p>
              </w:sdtContent>
            </w:sdt>
          </w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
 </pkg:package>
```

すでに説明したように、コンテンツ コンテキストは書式設定されたテキストと同様、追加のドキュメント パーツを必要としません。そのため、編集された .rels および document.xml パーツのみがここに含まれています。

document.xml 本文内に表示される **w:sdt** タグはコンテンツ コントロールを表します。コンテンツ コントロールの Office Open XML マークアップを生成すると、タグおよびドキュメント パーツ プロパティなど、いくつかの属性がこの例から削除されていることに気付くでしょう。次の内容を含む、不可欠な (そして、いくつかのベスト プラクティスの) 要素のみが維持されています。

- **エイリアス** は、Word の [コンテンツ コントロールのプロパティ] ダイアログ ボックスのタイトル プロパティです。 アドイン内からコントロールにバインドする場合、これは必須のプロパティです (アイテムの名前を表します)。

- 一意の **id** は必須のプロパティです。アドイン内からコントロールにバインドする場合、この ID は、該当する名前付きコンテンツ コントロールを識別するためにドキュメントのバインド機能で使用されるプロパティです。

- **外観** 属性は、よりクリーンな外観を実現するために、コントロール コンテナーを非表示にするために使用されます。 この機能は、w15 名前空間の使用でわかるように、Word 2013 で導入されました。 このプロパティが使用されるため、w15 名前空間が document.xml パーツの先頭で維持されます。

- **showingPlcHdr** 属性は、プレースホルダー コンテンツとしてコントロール内に含める既定のコンテンツ (この例のテキスト) を設定するオプションの設定です。 そのため、ユーザーがコントロール領域をクリックするか、タップすると、ユーザーが変更できる編集可能コンテンツのように動作する代わりに、コンテンツ全体が選択されます。

- **sdt** タグの前にある空の段落記号 ( **w:p/** ) は、コンテンツ コントロールを追加するために必要ありません (Word 文書のコントロールの上に垂直方向のスペースが追加されます)、コントロールが独自の段落に配置されていることを確認します。 コントロールに追加されるコンテンツのタイプと書式設定によっては、これが重要になる場合があります。

- コントロールにバインドする場合、バインドで複数の段落から構成されるリッチ コンテンツを受け入れるために、コントロールの既定のコンテンツ (**sdtContent** タグの内容) に少なくとも 1 つの完全な段落を含める必要があります。

> [!NOTE]
> このサンプルの **w:sdt** タグから削除されたドキュメント パーツ属性は、プレースホルダー コンテンツ情報を保存できるパッケージの別個のパーツ (Office Open XML パッケージの用語集ディレクトリに置かれているパーツ) を参照するために、コンテンツ コントロールに表示される場合があります。 ドキュメント パーツは Office Open XML パッケージ内の XML パーツ (つまりファイル) に使用される用語ですが、sdt プロパティで使用されるドキュメント パーツという用語は、文書パーツやドキュメント プロパティクイック パーツ (組み込みの XML データ バインド コントロールなど) を含む一部のコンテンツ タイプを記述するために使用される Word の用語と同じ用語を指します。 Office Open XML パッケージの用語集ディレクトリの下にパーツが表示される場合、挿入するコンテンツにこれらの機能が含まれるのであれば、そのパーツは維持する必要がある可能性があります。 アドインからバインドするために使用する一般的なコンテンツ コントロールの場合は、必要ありません。 パッケージから用語集パーツを削除する場合、w:sdt タグからもドキュメント パーツ属性を削除する必要があることを覚えておいてください。

次のセクションでは、Word アドインでバインドを作成し、使用する方法について考察します。

## <a name="insert-content-at-a-designated-location"></a>指定された場所にコンテンツを挿入する

Word 文書のアクティブな選択範囲にコンテンツを挿入する方法を既に確認しました。 ドキュメントの名前付きコンテンツ コントロールにバインドする場合、同じコンテンツ タイプをそのコントロールに挿入できます。

それでは、どのような場合にこの手法を使用するのでしょうか。

- データベースからドキュメントの一部にデータを入力するなど、テンプレートの指定の位置にコンテンツを追加したり、置換したりする場合

- デザイン要素オプションをユーザーに提供するなど、アクティブな選択範囲に挿入するコンテンツを置換するオプションが必要な場合。

- ユーザーがドキュメントに追加した情報に基づいて作業ウィンドウのフィールドにデータを入力するなど、アドインで使用するためにアクセスできるドキュメントに、ユーザーにデータを追加させたい場合

コード サンプル [Word-Add-in-JavaScript-AddPopulateBindings](https://github.com/OfficeDev/Word-Add-in-JavaScript-AddPopulateBindings) をダウンロードします。このサンプルは、コンテンツ コントロールを挿入し、それにバインドする方法と、バインドにデータを入力する方法の実際の例を提供します。

### <a name="add-and-bind-to-a-named-content-control"></a>名前付きコンテンツ コントロールを追加し、それにバインドする

続く JavaScript を確認し、これらの要件について考察してください。

- すでに説明したように、Word アドインからコントロールにバインドするには、リッチ テキスト コンテンツ コントロールを使用する必要があります。

- コンテンツ コントロールには名前が必要です (これは、[コンテンツ コントロールのプロパティ] ダイアログ ボックスの **[タイトル** ] フィールドで、Office Open XML マークアップの **エイリアス** タグに対応します)。 これにより、コードはバインドを配置する場所を識別します。

- 複数の名前付きコントロールを用意し、必要に応じてそれらにバインドできます。一意のコンテンツ コントロール名、一意のコンテンツ コントロール ID、一意のバインド ID を使用します。

```js
function addAndBindControl() {
    Office.context.document.bindings.addFromNamedItemAsync("MyContentControlTitle", "text", { id: 'myBinding' }, function (result) {
        if (result.status == "failed") {
            if (result.error.message == "The named item does not exist.")
                const myOOXMLRequest = new XMLHttpRequest();
                let myXML;
                myOOXMLRequest.open('GET', '../../Snippets_BindAndPopulate/ContentControl.xml', false);
                myOOXMLRequest.send();
                if (myOOXMLRequest.status === 200) {
                    myXML = myOOXMLRequest.responseText;
                }
                Office.context.document.setSelectedDataAsync(myXML, { coercionType: 'ooxml' }, function (result) {
                    Office.context.document.bindings.addFromNamedItemAsync("MyContentControlTitle", "text", { id: 'myBinding' });
                });
        }
    });
}
```

ここに示すコードでは、次の手順を実行します。

- [addFromNamedItemAsync](/javascript/api/office/office.bindings#office-office-bindings-addfromnameditemasync-member(1)) を使用して、名前付きコンテンツ コントロールにバインドしようとします。

  コードの実行時に名前付きコントロールがドキュメント内に既に存在する可能性があるアドインのシナリオがある場合は、まずこの手順を実行します。 たとえば、コントロールが事前に配置されているアドインで動作するように設計されているテンプレートにアドインが挿入されて、テンプレート共に保存される場合は、これを実行します。 また、アドインによって前に配置されたコントロールにバインドする必要がある場合にも、これを行う必要があります。

- メソッドの `addFromNamedItemAsync` 最初の呼び出しのコールバックは、結果の状態をチェックして、名前付き項目がドキュメントに存在しないためにバインドが失敗したかどうかを確認します (つまり、この例では MyContentControlTitle という名前のコンテンツ コントロール)。 その場合、コードはアクティブな選択ポイント (使用) にコントロールを `setSelectedDataAsync`追加し、それにバインドします。

> [!NOTE]
> 前のコードで説明したように、コンテンツ コントロールの名前を使用し、バインドを作成する場所が決まります。ただし、Office Open XML マークアップでは、コードにより、コンテンツ コントロールの名前と ID 属性の両方を使用し、ドキュメントにバインドが追加されます。

コードの実行後、アドインがバインドを作成したドキュメントのマークアップを調べると、各バインドに 2 つのパーツが表示されます。 バインドが追加されたコンテンツ コントロールのマークアップ (document.xml) には、属性 **w15:webExtensionLinked/** が表示されます。

In the document part named webExtensions1.xml, you'll see a list of the bindings you've created. それぞれは、該当するコントロールのバインド ID と ID 属性を使用して識別されます。 **appref** 属性はコンテンツ コントロール ID です。 **we:binding id="myBinding" type="text" appref="1382295294"/**.

> [!IMPORTANT]
> バインドに操作を行うときにバインドを追加する必要があります。コンテンツ コントロールを挿入するために Office Open XML にバインドのマークアップを含めないでください。そのマークアップを挿入する過程でバインドは削除されます。

### <a name="populate-a-binding"></a>バインドにデータを入力する

バインドにコンテンツを書き込むためのコードは、選択範囲にコンテンツを書き込むためのコードに似ています。

```js
function populateBinding(filename) {
  const myOOXMLRequest = new XMLHttpRequest();
  let myXML;
  myOOXMLRequest.open('GET', filename, false);
  myOOXMLRequest.send();
  if (myOOXMLRequest.status === 200) {
      myXML = myOOXMLRequest.responseText;
  }
  Office.select("bindings#myBinding").setDataAsync(myXML, { coercionType: 'ooxml' });
}
```

同様に `setSelectedDataAsync`、挿入するコンテンツと強制の種類を指定します。 バインドに書き込むための唯一の追加要件は、ID によりバインドを識別することです。 このコードで使用されるバインド ID (bindings#myBinding) が、前の関数でバインドが作成されたときに確立されたバインド ID (myBinding) にどのように対応しているかに注目してください。

> [!NOTE]
> バインドのコンテンツの最初のデータ入力であろうと、コンテンツの置換であろうと、必要なのは前のコードだけです。バインド位置にコンテンツの新しい部分を挿入するとき、そのバインドの既存のコンテンツが自動的に置換されます。前に参照したコード サンプル [Word-Add-in-JavaScript-AddPopulateBindings](https://github.com/OfficeDev/Word-Add-in-JavaScript-AddPopulateBindings) でこの例を確認してください。このコード サンプルでは、同じバインドに入力するために同じように使用できる 2 つの別個のコンテンツ サンプルが提供されます。

## <a name="add-objects-that-use-additional-office-open-xml-parts"></a>追加の Office Open XML パーツを使用するオブジェクトを追加する

多くのタイプのコンテンツで、Office Open XML パッケージの追加ドキュメント パーツが必要です。つまり、別のパーツの情報を参照するか、1 つまたは複数のパーツに格納され、document.xml で参照されるコンテンツ自体を参照します。

たとえば、次の内容を考慮します。

- 書式設定にスタイルを使用するコンテンツ (先ほどの図 2 のスタイル設定されたテキストまたは図 9 のスタイル設定された表など) には styles.xml パーツが必要です。

- 画像 (図 3 および図 4 のような画像など) では、1 つ (2 つの場合もある) の追加パーツにバイナリ画像が含まれます。

- SmartArt 図 (図 10 のような図) では、レイアウトとコンテンツを表現するために、複数の追加パーツが必要です。

- グラフ (図 11 のようなグラフなど) では、独自のリレーションシップ (.rels) パーツを含む、複数の追加パーツが必要です。

これらのすべてのコンテンツ タイプのマークアップの編集例については、以前に参照したコード サンプルの [Word アドインでの Open XML の読み込みと書き込みを参照してください](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/word-add-in-load-and-write-open-xml)。 バインディングを使用して現在の選択範囲にコンテンツを挿入したり、指定した場所にコンテンツを書き込んだりする場合は、前に表示した (および参照したコード サンプルで提供された) ものと同じ JavaScript コードを使用してこれらのコンテンツ タイプをすべて挿入できます。

サンプルを調べる前に、これらの各コンテンツ タイプを操作するためのいくつかのヒントを見てみましょう。

> [!IMPORTANT]
> document.xml で参照される追加パーツを維持する場合、document.xml.rels と、styles.xml や画像ファイルなど、維持する該当パーツのリレーションシップを維持する必要があります。

### <a name="working-with-styles"></a>スタイルを使用する

段落スタイルまたはテーブル スタイルを使用してコンテンツを書式設定する場合、直接書式設定されたテキストを使用して前の例で見たマークアップを編集するのと同じ方法が適用されます。 ただし、段落スタイルを操作するためのマークアップは非常に単純であるため、ここで説明する例を示します。

#### <a name="editing-the-markup-for-content-using-paragraph-styles"></a>段落スタイルを使用してコンテンツのマークアップを編集する

次のマークアップは、図 2 で示した、スタイル設定されたテキスト例の本文コンテンツを表します。

```XML
<w:body>
  <w:p>
    <w:pPr>
      <w:pStyle w:val="Heading1"/>
    </w:pPr>
    <w:r>
      <w:t>This text is formatted using the Heading 1 paragraph style.</w:t>
    </w:r>
  </w:p>
</w:body>
```

> [!NOTE]
> おわかりのように、document.xml の書式設定されたテキストのマークアップはスタイルを使用すると大幅にシンプルになります。スタイルには、スタイルを使用しなければ個々に参照しなければならない段落とフォントの書式設定がすべて含まれているからです。ただし、前述のように、さまざまな理由から、スタイルを使用することもあれば、直接の書式設定を使用することもあります。直接の書式設定を使用すれば、ユーザーのドキュメントの書式設定に関係なく、テキストの外見を指定できます。段落スタイルを使用すれば (特に、例にある見出し 1 など、組み込みの段落スタイル名)、テキストの書式設定をユーザーのドキュメントに合わせて自動的に調整できます。

スタイルの使用は、挿入するコンテンツのマークアップを読み取り、理解することがいかに重要であるかの良い例です。これは、別のドキュメント パーツがここで参照されていることは明示的ではないためです。 このマークアップにスタイル定義を含め、styles.xml パーツを含めない場合、document.xmlのスタイル情報は、そのスタイルがユーザーのドキュメントで使用されているかどうかに関係なく無視されます。

ただし、styles.xml part をよく見ると、この長いマークアップのほんの一部だけがアドインで使用するマークアップを編集するときに必要であることがわかります。

- styles.xml パーツには既定で複数の名前空間が含まれます。コンテンツに必要なスタイル情報のみを維持するのであれば、ほとんどの場合、**xmlns:w** 名前空間のみを維持するだけで問題ありません。

- スタイル パーツの上部にある **w:docDefaults** タグ コンテンツは、マークアップがアドイン経由で挿入されたときは無視されます。これは削除してもかまいません。

- styles.xml パーツの大部分のマークアップは、docDefaults の後に表示される **w:latentStyles** タグのものです。これは利用可能なすべてのスタイルのための情報 (Styles ウィンドウと Styles ギャラリーの外観属性など) を提供します。この情報もアドイン経由で挿入されたときに無視されます。そのため削除してもかまいません。

- 潜在的なスタイル情報を追うと、マークアップの生成元になったドキュメントで使用されている各スタイルの定義がわかります。これには、ドキュメントを新規作成するときに使用される既定のスタイルが含まれ、お使いのコンテンツには関係ない場合があります。お使いのコンテンツで使用されないスタイルの定義は削除できます。

   > [!NOTE]
   > 各組み込みの見出しスタイルには、同じ見出し形式の文字スタイル バージョンである Char スタイルが関連付けられます。 Unless you've applied the heading style as a character style, you can remove it. If the style is used as a character style, it appears in document.xml in a run properties tag ( **w:rPr** ) rather than a paragraph properties ( **w:pPr** ) tag. This should only be the case if you've applied the style to just part of a paragraph, but it can occur inadvertently if the style was incorrectly applied.

- コンテンツに組み込みスタイルを使用している場合、完全な定義を含める必要はありません。挿入時に型強制 Office Open XML でコンテンツにスタイルを適用する目的で、スタイル名、スタイル ID、そして少なくとも 1 つの書式設定属性だけを含める必要があります。

    ただし、完全なスタイル定義を含めるのがベスト プラクティスです (組み込みスタイルに対して既定の場合であっても)。スタイルがターゲット ドキュメントで既に使用されている場合は、styles.xml の内容に関係なく、コンテンツはスタイルの常駐定義を使用します。スタイルがターゲット ドキュメントでまだ使用されていない場合は、コンテンツはマークアップで指定されているスタイル定義を使用します。

そのため、たとえば、図 2 に示すサンプル テキストのstyles.xml パーツから保持する必要があるコンテンツは、見出し 1 スタイルを使用して書式設定されている唯一のコンテンツです。

> [!NOTE]
> この例では、見出し 1 スタイルの完全な Word 定義が維持されています。

```XML
<pkg:part pkg:name="/word/styles.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml">
  <pkg:xmlData>
    <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" >
      <w:style w:type="paragraph" w:styleId="Heading1">
        <w:name w:val="heading 1"/>
        <w:basedOn w:val="Normal"/>
        <w:next w:val="Normal"/>
        <w:link w:val="Heading1Char"/>
        <w:uiPriority w:val="9"/>
        <w:qFormat/>
        <w:pPr>
          <w:keepNext/>
          <w:keepLines/>
          <w:spacing w:before="240" w:after="0" w:line="259" w:lineRule="auto"/>
          <w:outlineLvl w:val="0"/>
        </w:pPr>
        <w:rPr>
          <w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/>
          <w:color w:val="2E74B5" w:themeColor="accent1" w:themeShade="BF"/>
          <w:sz w:val="32"/>
          <w:szCs w:val="32"/>
        </w:rPr>
      </w:style>
    </w:styles>
  </pkg:xmlData>
</pkg:part>
```

#### <a name="edit-the-markup-for-content-using-table-styles"></a>テーブル スタイルを使用してコンテンツのマークアップを編集する

コンテンツで表スタイルを使用するとき、段落スタイルの使用で説明したように、styles.xml の同じ相対パーツを使用します。つまり、コンテンツで使用しているスタイルの情報だけを維持すればよいのです。そこには、名前、ID、少なくとも 1 つの書式設定属性を含める必要があります。ただし、考えられるあらゆるユーザー シナリオに対処するために、完全なスタイル定義を含めるのが最善策です。

ただし、document.xml の表と styles.xml の表スタイル定義の両方のマークアップを見ると、段落スタイルを使用するときより、はるかに多くのマークアップがあるのがわかります。

- document.xml では、書式設定はスタイルに含まれている場合でも、セルごとに適用されます。表スタイルを使用してマークアップの量が減ることはありません。コンテンツに表スタイルを使用すると、複数の表を簡単に更新し、簡単に調整できるメリットがあります。

- styles.xml では、たった 1 つの表スタイルでもかなりの量のマークアップがあります。これは、表全体、見出し行、奇数と偶数の縞付き行および列 (別個)、最初の列など、表スタイルに表領域別のさまざまな書式設定属性が含まれるためです。

### <a name="work-with-images"></a>画像を操作する

イメージのマークアップには、イメージを記述するバイナリ データを含む少なくとも 1 つの部分への参照が含まれています。 複雑なイメージの場合、これは何百ページものマークアップで、編集することはできません。 バイナリパーツに触れる必要がないため、Visual Studio などの構造化エディターを使用している場合は折りたたむだけで、パッケージの残りの部分を簡単に確認および編集できます。

図 3 で前に参照したコード サンプル「 [Word アドインで Open XML を読み込んで書き込む](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/word-add-in-load-and-write-open-xml)」で使用できる単純な画像のマークアップ例を確認すると、document.xml内の画像のマークアップに、サイズと位置情報、およびバイナリ イメージ データを含む部分へのリレーションシップ参照が含まれていることがわかります。 その参照は、次のように、 **a:blip** タグに含まれています。

```XML
<a:blip r:embed="rId4" cstate="print">
```

リレーションシップ参照が明示的に使用される ( **r:embed="rID4"** ) ため、イメージをレンダリングするために関連する部分が必要であるため、Office Open XML パッケージにバイナリ データを含めない場合は、エラーが発生します。 これは、前述のstyles.xmlとは異なります。リレーションシップが明示的に参照されておらず、リレーションシップがコンテンツ自体の一部ではなく、コンテンツ (書式設定) に属性を提供する部分であるため、省略してもエラーはスローされません。

> [!NOTE]
> マークアップを確認すると、a:blip タグで使用される追加の名前空間に注目してください。 document.xmlでは、 **xlmns:** a 名前空間 (メインの drawingML 名前空間) が、document.xml パーツの上部ではなく、drawingML 参照の使用の開始時に動的に配置されていることがわかります。 However, the relationships namespace (r) must be retained where it appears at the start of document.xml. Check your picture markup for additional namespace requirements. Remember that you don't have to memorize which types of content require what namespaces, you can easily tell by reviewing the prefixes of the tags throughout document.xml.

### <a name="understanding-additional-image-parts-and-formatting"></a>追加の画像パーツと書式設定について理解する

図 4 の (ピクチャのスタイル設定に加えて) 調整済みの明るさやコントラストの設定を使用する画像のように、画像に Office ピクチャ書式設定効果を使用する場合、画像データの HD 書式コピーの二次バイナリ データ パーツが必要になる場合があります。この追加 HD 書式は、レイヤー効果を設定するために必要であり、その書式への参照は document.xml に次のように表示されます。

```XML
<a14:imgLayer r:embed="rId5">
```

Word アドイン コード サンプルの [Open XML の読み込みと書き込み](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/word-add-in-load-and-write-open-xml) で、図 4 に示す書式設定された画像に必要なマークアップ (レイヤー効果を使用) を参照してください。

### <a name="work-with-smartart-diagrams"></a>SmartArt ダイアグラムを操作する

SmartArt 図には 4 つの関連パーツがありますが、常に必要なのは 2 つだけです。 Word アドインのコード サンプルで、 [Open XML の読み込みと書き込みで](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/word-add-in-load-and-write-open-xml) SmartArt マークアップの例を確認できます。 最初に、各パーツの簡単な説明と、それらが必要な理由または不要な理由を見てみましょう。

> [!NOTE]
> コンテンツに複数の図が含まれている場合は、次に示すファイル名の '1' を置き換えて、連続して番号が付けられます。

- layout1.xml: このパーツは必須です。これには、レイアウトの外観と機能のマークアップ定義が含まれます。

- data1.xml: このパーツは必須です。これには、図のインスタンスで使用されるデータが含まれます。

- drawing1.xml: このパーツは必ずしも必要ではありませんが、個々の図形を直接書式設定するなど、図のインスタンス内の要素にカスタム書式を適用する場合は、それを保持する必要がある場合があります。

- colors1.xml: この部分は必要ありません。 色のスタイル定義が含まれますが、図の色は既定で、Office Open XML マークアップを保存する前に Word の SmartArt ツールの [デザイン] タブから適用する SmartArt 色スタイルに基づき、ターゲット ドキュメントのアクティブな書式設定テーマの色に合わせて調整されます。

- quickStyles1.xml: この部分は必要ありません。 ターゲット ドキュメントで利用できる、適用済みの SmartArt スタイルの定義を図が取得するため色パーツと同様に、このパーツは削除できます(つまり、パーツはターゲット ドキュメントの書式設定テーマに合わせて自動的に調整されます)。

> [!TIP]
> SmartArt layout1.xml ファイルは、マークアップをさらにトリミングできる場所の良い例ですが、余分な時間を費やすことはできません (パッケージ全体に対してこのような少量のマークアップが削除されるため)。 マークアップで削除できる最後の行まで削除する場合、**dgm:sampData** タグとそのコンテンツを削除します。 このサンプル データにより、SmartArt スタイル ギャラリーにおける図のサムネイル プレビューの表示が定義されます。 ただし、それを省略した場合、既定のサンプル データが使用されます。

document.xmlの SmartArt ダイアグラムのマークアップには、レイアウト、データ、色、クイック スタイルの各部分へのリレーションシップ ID 参照が含まれていることに注意してください。 これらのパーツとそのリレーションシップ定義を削除すると、色とスタイル パーツに対するdocument.xml内の参照を削除できます (また、これらのリレーションシップを削除するので、これを行うのがベスト プラクティスであることは確かですが)、図面をドキュメントに挿入する必要がないため、残してもエラーは発生しません。 **dgm:relIds** タグのdocument.xmlでこれらの参照を見つけます。 この手順を実行するかどうかに関係なく、必要なレイアウトとデータパーツのリレーションシップ ID 参照を保持します。

### <a name="work-with-charts"></a>グラフを操作する

SmartArt 図と同様に、グラフにはいくつかの追加のパーツが含まれます。 ただし、グラフの設定は SmartArt とは少し異なり、グラフには独自のリレーションシップ ファイルがあります。 次に、グラフに必要なリムーバブル ドキュメント パーツの説明を示します。

> [!NOTE]
> SmartArt 図と同様に、コンテンツに複数のグラフが含まれている場合は、連続して番号が付けられます。次に示すファイル名の '1' が置き換えられます。

- document.xml.rels には、グラフ (chart1.xml) を表すデータを含む必須パーツへの参照があります。

- また、chart1.xml.rels など、Office Open XML パッケージのグラフ別のリレーションシップ ファイルがあります。

    chart1.xml.rels では 3 つのファイルが参照されますが、1 つだけが必須です。この中にバイナリ Excel ワークブック データ (必須) と、削除できる色およびスタイル パーツ (colors1.xml と styles1.xml) があります。

Word でネイティブで作成し、編集できるグラフは Excel グラフです。そのデータは、Office Open XML パッケージにバイナリ データとして組み込まれている Excel ワークシートで管理されます。画像のバイナリ データ パーツのように、この Excel バイナリ データは必須ですが、このパーツでは何も編集しません。そのため、エディターでパーツを折りたためば、手動で最後までスクロールしなくても Office Open XML パッケージの残りを調べることができます。

ただし、SmartArt と同様に、色およびスタイル パーツを削除できます。 グラフの書式設定に使用できるグラフのスタイルと色のスタイルを使用した場合、グラフはコピー先のドキュメントに挿入されるときに、該当する書式を自動的に適用します。

Word アドインのコード サンプルの [Open XML の読み込みと書き込みを](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/word-add-in-load-and-write-open-xml) 行う図 11 に示すグラフの例については、編集済みのマークアップを参照してください。

## <a name="edit-the-office-open-xml-for-use-in-your-task-pane-add-in"></a>作業ウィンドウ アドインで使用する Office Open XML を編集する

マークアップのコンテンツの識別方法と編集方法はすでに確認しました。 ドキュメント用に生成された大規模な Office Open XML パッケージを見て、タスクがまだ難しいと思われる場合は、そのパッケージをすばやく編集するのに役立つ推奨される手順の簡単な概要を次に示します。

> [!NOTE]
> パッケージのすべての .rels パーツをマップとして使用し、削除できるドキュメント パーツをすばやく確認できることを覚えておいてください。

1. 単層化された XML ファイルを Visual Studio で開き、Ctrl+K、Ctrl+D を押し、ファイルを書式設定します。次に、左側の折りたたみ/展開ボタンを使用し、削除する必要があることがわかっているパーツを折りたたみます。必要であるが、編集する必要はない長いパーツを折りたたむこともできます (画像ファイルの base64 バイナリ データなど)。このようにすると、マークアップが視覚的に楽に速く読めるようになります。

1. ドキュメント パッケージのいくつかのパーツは、アドインに使用する Office Open XML マークアップを準備しているときにほとんどの場合は削除できます。 まずは、そのようなパーツ (およびそれらに関連付けられているリレーションシップの定義) を削除すれば、パッケージのサイズをすぐに大幅に減らすことができます。 そのようなパーツには、theme1、fontTable、settings、webSettings、サムネイル、コア プロパティ ファイルとアドイン プロパティ ファイルの両方、任意の `taskpane`、`webExtension` パーツなどがあります。

1. 必要のない脚注、ヘッダー、フッターなど、自分のコンテンツに関連しないパーツを削除できます。ここでも、関連付けられているリレーションシップを必ず削除してください。

1. document.xml.rels パーツをレビューし、画像ファイル、スタイル パーツ、SmartArt 図パーツなど、そのパーツで参照されるファイルが自分のコンテンツに必要かどうかを確認します。自分のコンテンツに必要のないパーツがあれば、そのリレーションシップを削除し、関連パーツも削除されていることを確認します。自分のコンテンツには document.xml.rels で参照されるドキュメント パーツがいずれも必要でない場合、そのファイルを削除することもできます。

1. 自分のコンテンツに .rels パーツ (chart#.xml.rels など) を追加している場合、それをレビューし、そこで参照されているその他のパーツで、削除できるパーツがないかどうかを確認し (グラフのクイック スタイルなど)、そのファイルからのリレーションシップと関連パーツの両方を削除します。

1. パーツで参照されていない名前空間、コンテンツにセクション区切りが含まれていない場合はセクション プロパティ、挿入するコンテンツに関連しないマークアップを削除するには、document.xmlを編集します。 図形またはテキスト ボックスを挿入する場合、広範囲のフォールバック マークアップを削除することもできます。

1. スタイル パーツなど、コンテンツに影響を与えずに重要なマークアップを削除できることがわかっている必須の追加パーツを編集します。

前の 7 つの手順を終えると、コンテンツにもよりますが、削除できるマークアップの約 90 から 100% を削除していることになります。ほとんどの場合、削除作業はこれで終了でしょう。

ここで終了するか、削除できる最後の行まで見つけるためにコンテンツをさらに深く調査するかにかかわらず、前に参照したコード サンプル [Word-Add-in-Get-Set-EditOpen-XML](https://github.com/OfficeDev/Word-Add-in-Get-Set-EditOpen-XML) をスクラッチ パッドとして使用し、編集したマークアップをすばやくかつ簡単にテストできることを覚えておいてください。

> [!TIP]
> 開発中に既存のソリューションで Office Open XML スニペットを更新する場合は、ソリューションを再実行する前に一時インターネット ファイルをクリアして、コードで使用されている Office Open XML を更新します。 Markup that's included in your solution in XML files is cached on your computer. You can, of course, clear temporary Internet files from your default web browser. インターネット オプションにアクセスし、Visual Studio 2019 内からこれらの設定を削除するには、[ **デバッグ** ] メニューの [ **オプション**] を選択します。 次に、[**環境**] で、[**Web ブラウザー**] を選択し、[**Internet Explorer オプション**] を選択します。

## <a name="create-an-add-in-for-both-template-and-standalone-use"></a>テンプレートとスタンドアロンの両方で使用するアドインを作成する

このトピックでは、アドインで Office Open XML で実行できる操作の例をいくつか紹介しました。Office Open XML 強制型を使用してドキュメントに挿入できる豊富なコンテンツ タイプの例を、選択範囲または指定された (バインドされた) 場所に挿入するための JavaScript メソッドと共に、さまざまな豊富なコンテンツ タイプの例を見てきました。

そのため、アドインをスタンドアロンで使用する場合 (つまり、ストアまたは独自のサーバーの場所から挿入される) と、アドインを操作するように設計された事前に作成されたテンプレートで使用する場合は、他に何を知る必要がありますか? 必要なことはすべて知っているというのが答えかもしれません。

特定のコンテンツ タイプとそれを挿入するためのメソッドのマークアップは、アドインがスタンドアロンに設計されているか、テンプレートを操作するように設計されているかに関係なく同じです。 自分のアドインで作動するように設計されたテンプレートを使用する場合、参照コンテンツがドキュメントにすでに存在するシナリオを構成するコールバックを必ず JavaScript に含めます (セクション「[指定コンテンツ コントロールを追加し、それにバインドする](#add-and-bind-to-a-named-content-control)」のバインド例)。

ユーザーがドキュメントを作成したときにアドインがテンプレートに常駐しているか、アドインがテンプレートを挿入するかに関係なく、自分のアドインでテンプレートを使用するときは、API のその他の要素を組み込み、さらに強固でインタラクティブな体験を作ることもできます。たとえば、テンプレート固有のオプションをユーザーに提供するために、customXML パーツに識別データを含め、それを使用し、テンプレート タイプを決定できます。アドインにおける customXML の動作に関する詳細は、以下のその他の技術情報を参照してください。

## <a name="see-also"></a>関連項目

- [Office の JavaScript API](../reference/javascript-api-for-office.md)
- Open XML: [Standard ECMA-376: Office Open XML ファイル形式](https://www.ecma-international.org/publications/standards/Ecma-376.htm)に関する完全な言語リファレンスと関連ドキュメント
- [Office JavaScript API の探索: データ バインディングとカスタム XML パーツ](/archive/msdn-magazine/2013/april/microsoft-office-exploring-the-javascript-api-for-office-data-binding-and-custom-xml-parts)
