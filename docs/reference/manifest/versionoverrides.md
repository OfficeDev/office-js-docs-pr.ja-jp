---
title: マニフェスト ファイルの VersionOverrides 要素
description: アドイン マニフェスト (XML) ファイルOffice VersionOverrides 要素のリファレンス ドキュメント。
ms.date: 01/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: 657bdebbc88993badd9d0e60946239edd55d5533
ms.sourcegitcommit: 9b0e70bb296a84adfaea0d6fee54916be9e13031
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/14/2022
ms.locfileid: "62042148"
---
# <a name="versionoverrides-element"></a>VersionOverrides 要素

この要素には、基本マニフェストでサポートされていない機能の情報が含まれています。 子マークアップは、基本マニフェスト (または親 VersionOverrides) のマークアップの一部 **を上書きする可能性があります**。 **VersionOverrides は** 、マニフェストの [ルート OfficeApp](officeapp.md) 要素または親 **VersionOverrides** 要素のいずれかの子要素です。 この要素はマニフェスト スキーマ v1.1 以降でサポートされますが、個別の VersionOverrides スキーマで定義されます。

詳細については、「マニフェストの [バージョンオーバーライド」を参照してください](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)。

## <a name="attributes"></a>属性

|  属性  |  必須  |  説明  |
|:-----|:-----|:-----|
|  **xmlns**       |  はい  |  VersionOverrides スキーマ名前空間。 許可される値は、この要素の `<VersionOverrides>` **xsi:type** 値と親要素の **xsi:type** 値によって異 `<OfficeApp>` なります。 以下の [名前空間の値を参照](#namespace-values) してください。|
|  **xsi:type**  |  はい  | スキーマのバージョン。現時点では、`VersionOverridesV1_0` および `VersionOverridesV1_1` のみが有効な値になります。 |

### <a name="namespace-values"></a>名前空間の値

ルート要素の **xsi:type** 値に応じて **、xmlns** 属性の必要な値を次に示 `<OfficeApp>` します。

- **TaskPaneApp は** VersionOverrides のバージョン 1.0 のみをサポートし **、xmlns は** `http://schemas.microsoft.com/office/taskpaneappversionoverrides` .
- **ContentApp** は VersionOverrides のバージョン 1.0 のみをサポートし **、xmlns は** `http://schemas.microsoft.com/office/contentappversionoverrides` .
- **MailApp** は VersionOverrides のバージョン 1.0 と 1.1 をサポートしています。 **したがって、xmlns** の値は、この要素の `<VersionOverrides>` **xsi:type** 値によって異なります。
  - **xsi:type がである** 場合 `VersionOverridesV1_0` は **、xmlns を** 指定する必要があります `http://schemas.microsoft.com/office/mailappversionoverrides` 。
  - **xsi:type がである** 場合 `VersionOverridesV1_1` は **、xmlns を** 指定する必要があります `http://schemas.microsoft.com/office/mailappversionoverrides/1.1` 。

> [!NOTE]
> 現在のところ、Outlook 2016以降は VersionOverrides v1.1 スキーマと型をサポート `VersionOverridesV1_1` しています。

## <a name="variant-schemas"></a>バリアント スキーマ

可能な **xmlns** 値ごとに異なるスキーマが用意されています。そのため、それぞれ個別の参照ページがあります。

- [VersionOverrides 1.0 TaskPane](versionoverrides-1-0-taskpane.md)
- [VersionOverrides 1.0 コンテンツ](versionoverrides-1-0-content.md)
- [VersionOverrides 1.0 Mail](versionoverrides-1-0-mail.md)
- [VersionOverrides 1.1 Mail](versionoverrides-1-1-mail.md)
