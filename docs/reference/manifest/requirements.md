---
title: マニフェスト ファイルの Requirements 要素
description: Requirements 要素は、最小要件セットを指定し、Office アドインを Office によってアクティブ化するか、基本マニフェスト設定を上書きする必要があるメソッドを指定します。
ms.date: 01/22/2022
ms.localizationpriority: medium
ms.openlocfilehash: 85dcd08f3bfcffe34c4c479608f25ea0c2b6a134
ms.sourcegitcommit: ae3a09d905beb4305a6ffcbc7051ad70745f79f9
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/26/2022
ms.locfileid: "62222284"
---
# <a name="requirements-element"></a>Requirements 要素

この要素の意味は、基本マニフェスト []](#in-the-base-manifest)で使用するか [**、VersionOverrides**](#as-a-child-of-a-versionoverrides-element)要素の子として使用するかによって異なります。

> [!TIP]
> この要素を使用する前に、「ホストと API の要件[Officeを指定する」を理解してください。](../../develop/specify-office-hosts-and-api-requirements.md)

## <a name="in-the-base-manifest"></a>基本マニフェストで

基本マニフェスト ([つまり、OfficeApp](officeapp.md)の直接の子) で使用する場合 **、Requirements** 要素は、Office アドインを Office でアクティブ化する必要がある Office JavaScript API 要件 [(要件](../../develop/office-versions-and-requirement-sets.md#specify-office-applications-and-requirement-sets)セットまたはメソッド) の最小セットを指定します。 指定されたメソッドと要件セットをサポートしていない Office バージョンとプラットフォーム (Windows、Mac、Web、iOS、iPad など) の組み合わせでは、アドインはアクティブ化されません。

**アドインの種類:** 作業ウィンドウ, メール

## <a name="as-a-child-of-a-versionoverrides-element"></a>VersionOverrides 要素の子として

[VersionOverrides](versionoverrides.md)の子として使用する場合は、基本マニフェスト設定を上書きする **VersionOverrides** 要素の設定のために、Office バージョンとプラットフォーム (Windows、Mac、Web、iOS、iPad など) でサポートする必要がある Office JavaScript API 要件 [(要件](../../develop/office-versions-and-requirement-sets.md#specify-office-applications-and-requirement-sets)セットやメソッド) の最小セットを指定します。  を有効にします。

基本マニフェストで要件 A を指定し **、VersionOverrides** 内で要件 B を指定するアドインを検討してください。 

- プラットフォームと Office バージョンが A をサポートしない場合、アドインはアクティブ化されません。Office はマニフェストの **VersionOverrides** セクションを解析しません。 
- A と B の両方がサポートされている場合は、アドインがアクティブ化され **、VersionOverrides** のすべてのマークアップが有効になります。 
- A がサポートされているが B がサポートされていない場合は、アドインがアクティブ化され **、VersionOverrides** のマークアップの一部が有効になります。 具体的には、基本マニフェスト要素を **オーバーライドしない VersionOverrides** の子要素が有効になります。 たとえば **、WebApplicationInfo** 要素または **EquivalentAddins が** 有効になります。 ただし、Hosts などの基本マニフェスト要素をオーバーライドする **VersionOverrides** のすべての子要素は有効ではありません。 代わりに、Officeオーバーライドされた基本マニフェスト マークアップの値を使用します。 

**アドインの種類:** 作業ウィンドウ, メール

**次の VersionOverrides スキーマでのみ有効です**。

- 作業ウィンドウ 1.0
- メール 1.0
- メール 1.1

詳細については、「マニフェストの [バージョンオーバーライド」を参照してください](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)。

**次の要件セットに関連付けられている**。

- [AddinCommands 1.1](../requirement-sets/add-in-commands-requirement-sets.md) 親 **VersionOverrides が** Taskpane 1.0 と入力されている場合。
- 親 **VersionOverrides** が Mail 1.0 と入力されている場合のメールボックス [1.3。](../../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md)
- 親 **VersionOverrides** が Mail 1.1 と入力されている場合のメールボックス [1.5。](../../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md)

### <a name="remarks"></a>注釈

Requirements **要素** は、基本マニフェストの Requirements で指定されていない追加の要件を指定しない場合 **、VersionOverrides** では目的を提供しません。  Officeプラットフォームが基本マニフェストの要件をサポートしない場合、アドインはアクティブ化されません **。VersionOverrides** 要素は解析されません。 このため、次の両方の条件が満たされている場合にのみ **、VersionOverrides** で Requirements 要素を使用する必要があります。

- アドインには **、VersionOverrides** の構成 (アドイン コマンドなど) で実装され、基本マニフェストの **Requirements** 要素で指定されていないメソッドまたは要件セットが必要な追加機能があります。
- アドインは便利で、追加機能に必要な要件をサポートしないプラットフォームと Office バージョンの組み合わせでも、追加機能なしでアクティブ化する必要があります。

> [!TIP]
> VersionOverrides **内の** 基本マニフェストから **Requirement 要素を繰り返し使用しない**。 これを行う場合、何の影響も及び **、VersionOverrides** 内の **Requirements** 要素の目的に関して誤解を招く可能性があります。

> [!WARNING]
> **VersionOverrides** で **Requirements** 要素を使用する前に、要件をサポートしないプラットフォームとバージョンの組み合わせでは、要件を必要としない機能を呼び出すアドインコマンドもインストールされません。 たとえば、2 つのカスタム リボン ボタンを持つアドインを検討します。 そのうちの 1 つはOffice **ExcelApi 1.4** (以降) で使用できる JavaScript API を呼び出します。 その他の呼び出し API は **、ExcelApi 1.9** (以降) でのみ使用できます。 **VersionOverrides** に **ExcelApi 1.9** の要件を設定すると、どちらのボタンもリボンに表示されません。 このシナリオのより良い戦略は、「ランタイム チェックでメソッドと要件セットのサポートをチェックする」で説明されている [手法を使用する方法です](../../develop/specify-office-hosts-and-api-requirements.md#runtime-checks-for-method-and-requirement-set-support)。 2 番目のボタンによって呼び出されるコードは、 `isSetSupported` **まず ExcelApi 1.9** のサポートを確認するために使用します。 サポートされていない場合、このコードは、アドインのこの機能がバージョンの Office で使用できないというメッセージをユーザーに提供します。 

> [!NOTE]
> Mail アドインでは **、VersionOverrides** 1.1 を **VersionOverrides** 1.0 内に入れ子にすることができます。 Office常に、プラットフォームおよびバージョンでサポートされている最高バージョン **の VersionOverrides** をOfficeします。

## <a name="syntax"></a>構文

```XML
<Requirements>
   ...
</Requirements>
```

## <a name="contained-in"></a>含まれる場所

[OfficeApp](officeapp.md) 
[VersionOverrides](versionoverrides.md)

## <a name="can-contain"></a>含めることができるもの

|要素|コンテンツ|メール|TaskPane|
|:-----|:-----|:-----|:-----|
|[Sets](sets.md)|x|x|x|
|[メソッド](methods.md)|x||x|

## <a name="see-also"></a>関連項目

利用できる要件セットの詳細については、「[Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)」を参照してください。
