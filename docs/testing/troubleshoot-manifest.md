---
title: Office アドインのマニフェストを検証する
description: XML スキーマおよび他のツールを使用して、Officeアドインのマニフェストを検証する方法について説明します。
ms.date: 03/24/2022
ms.localizationpriority: medium
ms.openlocfilehash: 02273647a2e1f639aec7adb260453a811035d7e4
ms.sourcegitcommit: 64942cdd79d7976a0291c75463d01cb33a8327d8
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/25/2022
ms.locfileid: "64404722"
---
# <a name="validate-an-office-add-ins-manifest"></a>Office アドインのマニフェストを検証する

アドインのマニフェスト ファイルを検証して、それが正しくて完全であることを確認します。 検証を行うと、アドインをサイドロードするときに「アドイン マニフェストが無効です」というエラーが発生している問題も特定することができます。 この記事では、マニフェスト ファイルを検証するための複数の方法について説明します。

> [!NOTE]
> ランタイム ログを使用してアドインのマニフェストでの問題をトラブルシューティングする方法の詳細については、「[ランタイム ログを使用してアドインをデバッグする](runtime-logging.md)」を参照してください。

## <a name="validate-your-manifest-with-the-yeoman-generator-for-office-add-ins"></a>Office アドイン用の Yeoman ジェネレーターでマニフェストを検証する

[Office アドイン用の Yeoman ジェネレーター](../develop/yeoman-generator-overview.md)を使用してアドインを作成した場合は、それを使用してプロジェクトのマニフェスト ファイルを検証することもできます。 プロジェクトのルート ディレクトリから次のコマンドを実行します。

```command&nbsp;line
npm run validate
```

![コマンド ラインで実行されている Yo Office検証ツールを示し、検証が渡された結果を生成するアニメーション GIF。](../images/yo-office-validator.gif)

> [!NOTE]
> この機能にアクセスするには、Office アドイン バージョン 1.1.17 以降の [Yeoman](../develop/yeoman-generator-overview.md) ジェネレーターを使用してアドイン プロジェクトを作成する必要があります。

[!INCLUDE [validate also runs Office Store validation](../includes/office-store-validate.md)]

## <a name="validate-your-manifest-with-office-addin-manifest"></a>office-addin-manifest を使用してマニフェストを検証する

[Office アドイン用の Yeoman ジェネレーター](../develop/yeoman-generator-overview.md)を使用せずアドインを作成した場合は、[office-addin-manifest](https://www.npmjs.com/package/office-addin-manifest) を使用してマニフェストを検証することもできます。

1. [Node.js](https://nodejs.org/download/) をインストールします。

1. コマンド プロンプトを開き、次のコマンドで検証ツールをインストールします。

    ```command&nbsp;line
    npm install -g office-addin-manifest
    ```

1. プロジェクトのルート ディレクトリ *で次のコマンドを実行します*。

    ```command&nbsp;line
    npm run validate
    ```

    > [!NOTE]
    > このコマンドが使用できない場合や機能しない場合は、代わりに次のコマンドを実行して、office-addin-manifest ツールの最新バージョンを強制的に使用します ( `MANIFEST_FILE` マニフェスト ファイルの名前に置き換えてください)。
    >
    > ```command&nbsp;line
    > npx office-addin-manifest validate MANIFEST_FILE
    > ```

[!INCLUDE [validate also runs Office Store validation](../includes/office-store-validate.md)]

このコマンドで問題が発生した場合は、次の手順を実行します ( `MANIFEST_FILE` マニフェスト ファイルの名前に置き換えてください)。

```command&nbsp;line
npx office-addin-manifest validate -p MANIFEST_FILE
```

## <a name="validate-your-manifest-against-the-xml-schema"></a>XML スキーマと比較してマニフェストを検証する

マニフェストは、[XML スキーマ定義 (XSD)](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8) ファイルと比較して検証することができます。 マニフェスト ファイルが、使用している要素のすべての名前空間を含む、正しいスキーマに従っていることを確認します。 他のマニフェストのサンプルから要素をコピーした場合は、**適切な名前空間が含まれている** ことも再確認します。 XML スキーマの検証ツールを使用して、この検証を実行できます。

### <a name="to-use-a-command-line-xml-schema-validation-tool-to-validate-your-manifest"></a>コマンド ライン XML スキーマ検証ツールを使用してマニフェストを検証するには

1. [tar](https://www.gnu.org/software/tar/) および [libxml](http://xmlsoft.org/FAQ.html) をまだインストールしていない場合はインストールします。

1. 次のコマンドを実行します。`XSD_FILE` をマニフェスト XSD ファイルへのパスに置き換え、`XML_FILE` をマニフェスト XML ファイルへのパスに置き換えます。

    ```command&nbsp;line
    xmllint --noout --schema XSD_FILE XML_FILE
    ```

## <a name="see-also"></a>関連項目

- [Office アドインの XML マニフェスト](../develop/add-in-manifests.md)
- [Office のキャッシュをクリアする](clear-cache.md)
- [ランタイム ログを使用してアドインをデバッグする](runtime-logging.md)
- [テスト用に Office アドインをサイドロードする](sideload-office-add-ins-for-testing.md)
- [Internet Explorer の開発者ツールを使用してアドインをデバッグする](debug-add-ins-using-f12-tools-ie.md)
- [Edge レガシー用の開発者ツールを使用してアドインをデバッグする](debug-add-ins-using-devtools-edge-legacy.md)
- [Microsoft Edge (Chromium ベース)で開発者ツールを使用してアドインをデバッグする](debug-add-ins-using-devtools-edge-chromium.md)
