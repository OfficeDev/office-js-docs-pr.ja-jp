---
title: Office アドインのマニフェストを検証する
description: XML スキーマやその他のツールを使用して Office アドインのマニフェストを検証する方法について説明します。
ms.date: 12/31/2019
localization_priority: Normal
ms.openlocfilehash: 9cd1c353d6f73decb5e39df96cf66da5912b8f9c
ms.sourcegitcommit: a0262ea40cd23f221e69bcb0223110f011265d13
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42688809"
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

![コマンドラインから Yo Office 検証コントロールが実行され、検証の成功結果が生成されたアニメーション gif](../images/yo-office-validator.gif)

> [!NOTE]
> この機能にアクセスするには、アドイン プロジェクトが [Office アドイン用の Yeoman ジェネレーター](https://www.npmjs.com/package/generator-office) バージョン 1.1.17 以降を使用して作成されている必要があります。

## <a name="validate-your-manifest-with-office-addin-manifest"></a>office-addin-manifest を使用してマニフェストを検証する

[Office アドイン用の Yeoman ジェネレーター](https://www.npmjs.com/package/generator-office)を使用せずアドインを作成した場合は、[office-addin-manifest](https://www.npmjs.com/package/office-addin-manifest) を使用してマニフェストを検証することもできます。

1. [Node.js](https://nodejs.org/download/) をインストールします。

2. プロジェクトのルート ディレクトリから次のコマンドを実行します。 `MANIFEST_FILE` をマニフェスト ファイルの名前に置き換えます。

    ```command&nbsp;line
    npx office-addin-manifest validate MANIFEST_FILE
    ```

    > [!NOTE]
    > このコマンドを実行すると、「コマンドの構文が無効です」というエラーメッセージが表示されます。 (`validate` コマンドが認識されないため)、次のコマンドを実行してマニフェストを検証します (`MANIFEST_FILE` をマニフェスト ファイル名で置き換えます)。 
    >
    > `npx --ignore-existing office-addin-manifest validate MANIFEST_FILE`

## <a name="validate-your-manifest-against-the-xml-schema"></a>XML スキーマと比較してマニフェストを検証する

マニフェストは、[XML スキーマ定義 (XSD)](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8) ファイルと比較して検証することができます。 マニフェスト ファイルが、使用している要素のすべての名前空間を含む、正しいスキーマに従っていることを確認します。 他のマニフェストのサンプルから要素をコピーした場合は、**適切な名前空間が含まれている**ことも再確認します。 XML スキーマの検証ツールを使用して、この検証を実行できます。

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