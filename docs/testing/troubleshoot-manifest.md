---
title: Office アドインのマニフェストを検証する
description: XML スキーマおよび他のツールを使用して、Officeアドインのマニフェストを検証する方法について説明します。
ms.date: 09/18/2020
localization_priority: Normal
ms.openlocfilehash: 8ac543719011339b3f22a6fb77f01c0e9ed6d930
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/23/2021
ms.locfileid: "53075538"
---
# <a name="validate-an-office-add-ins-manifest"></a>Office アドインのマニフェストを検証する

アドインのマニフェスト ファイルを検証して、それが正しくて完全であることを確認します。 検証を行うと、アドインをサイドロードするときに「アドイン マニフェストが無効です」というエラーが発生している問題も特定することができます。 この記事では、マニフェスト ファイルを検証するための複数の方法について説明します。

> [!NOTE]
> ランタイム ログを使用してアドインのマニフェストでの問題をトラブルシューティングする方法の詳細については、「[ランタイム ログを使用してアドインをデバッグする](runtime-logging.md)」を参照してください。

## <a name="validate-your-manifest-with-the-yeoman-generator-for-office-add-ins"></a>Office アドイン用の Yeoman ジェネレーターでマニフェストを検証する

[Office アドイン用の Yeoman ジェネレーター](https://www.npmjs.com/package/generator-office)を使用してアドインを作成した場合は、それを使用してプロジェクトのマニフェスト ファイルを検証することもできます。 プロジェクトのルート ディレクトリから次のコマンドを実行します。

```command&nbsp;line
npm run validate
```

![コマンド ラインで実行され、検証Office渡された結果を生成する、Yo の値を示すアニメーション GIF。](../images/yo-office-validator.gif)

> [!NOTE]
> この機能にアクセスするには、アドイン プロジェクトが [Office アドイン用の Yeoman ジェネレーター](https://www.npmjs.com/package/generator-office) バージョン 1.1.17 以降を使用して作成されている必要があります。

## <a name="validate-your-manifest-with-office-addin-manifest"></a>office-addin-manifest を使用してマニフェストを検証する

[Office アドイン用の Yeoman ジェネレーター](https://www.npmjs.com/package/generator-office)を使用せずアドインを作成した場合は、[office-addin-manifest](https://www.npmjs.com/package/office-addin-manifest) を使用してマニフェストを検証することもできます。

1. [Node.js](https://nodejs.org/download/) をインストールします。

2. コマンド プロンプトを開き、次のコマンドで検証ツールをインストールします。

    ```command&nbsp;line
    npm install -g office-addin-manifest
    ```

3. プロジェクトのルート ディレクトリ *で次のコマンドを実行します*。

    ```command&nbsp;line
    npm run validate
    ```

    > [!NOTE]
    > このコマンドが使用できない場合や機能していない場合は、代わりに次のコマンドを実行して、office-addin-manifest ツールの最新バージョンを強制的に使用します (マニフェスト ファイルの名前に置き `MANIFEST_FILE` 換えてください)。
    >
    > ```command&nbsp;line
    > npx --ignore-existing office-addin-manifest validate MANIFEST_FILE
    > ```

## <a name="validate-your-manifest-against-the-xml-schema"></a>XML スキーマと比較してマニフェストを検証する

マニフェストは、[XML スキーマ定義 (XSD)](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8) ファイルと比較して検証することができます。 マニフェスト ファイルが、使用している要素のすべての名前空間を含む、正しいスキーマに従っていることを確認します。 他のマニフェストのサンプルから要素をコピーした場合は、**適切な名前空間が含まれている** ことも再確認します。 XML スキーマの検証ツールを使用して、この検証を実行できます。

### <a name="to-use-a-command-line-xml-schema-validation-tool-to-validate-your-manifest"></a>コマンド ライン XML スキーマ検証ツールを使用してマニフェストを検証するには

1. [tar](https://www.gnu.org/software/tar/) および [libxml](http://xmlsoft.org/FAQ.html) をまだインストールしていない場合はインストールします。

2. 次のコマンドを実行します。`XSD_FILE` をマニフェスト XSD ファイルへのパスに置き換え、`XML_FILE` をマニフェスト XML ファイルへのパスに置き換えます。
    
    ```command&nbsp;line
    xmllint --noout --schema XSD_FILE XML_FILE
    ```

## <a name="see-also"></a>関連項目

- [Office アドインの XML マニフェスト](../develop/add-in-manifests.md)
- [Office のキャッシュをクリアする](clear-cache.md)
- [ランタイム ログを使用してアドインをデバッグする](runtime-logging.md)
- [テスト用に Office アドインをサイドロードする](sideload-office-add-ins-for-testing.md)
- [Office アドインをデバッグする](debug-add-ins-using-f12-developer-tools-on-windows-10.md)
