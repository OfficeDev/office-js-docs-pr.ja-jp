---
ms.date: 01/29/2019
description: Excel でカスタム関数を使用してユーザーを認証します。
title: カスタム関数の認証
ms.openlocfilehash: 260f15c39758b82a2145474f543c3c9ff5edd132
ms.sourcegitcommit: 70ef38a290c18a1d1a380fd02b263470207a5dc6
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/15/2019
ms.locfileid: "30052736"
---
# <a name="authentication"></a>認証

一部のシナリオでは、カスタム関数は、保護されたリソースにアクセスするためにユーザーを認証する必要があります。 カスタム関数は、特定の認証方法を必要としませんが、カスタム関数は、アドインの作業ウィンドウや他の UI 要素とは別のランタイムで実行されることに注意してください。 そのため、 `AsyncStorage`オブジェクトとダイアログ API を使用して、2つのランタイム間でデータをやり取りする必要があります。
  
## <a name="asyncstorage-object"></a>asyncstorage オブジェクト

カスタム関数ランタイムには、通常`localStorage` 、データを格納するグローバルウィンドウで使用できるオブジェクトがありません。 代わりに、データを設定して取得するために、 [officeruntime](https://docs.microsoft.com/javascript/api/office-runtime/officeruntime.asyncstorage)を使用して、カスタム関数と作業ウィンドウ間でデータを共有する必要があります。 

さらに、を使用`AsyncStorage`するメリットがあります。セキュリティで保護されたサンドボックス環境を使用して、他のアドインがデータにアクセスできないようにします。  

### <a name="suggested-usage"></a>推奨される使用法

作業ウィンドウまたはカスタム関数から認証する必要がある場合は、asyncstorage でアクセストークンが既に取得されているかどうかを確認します。 それ以外の場合は、ダイアログ API を使用してユーザーを認証し、アクセストークンを取得してから、トークンを asyncstorage に保存しておきます。

## <a name="dialog-api"></a>ダイアログ API

トークンが存在しない場合は、ダイアログ API を使用して、ユーザーにサインインを要求する必要があります。 ユーザーが資格情報を入力すると、作成されたアクセストークン`AsyncStorage`がに保存されます。

> [!NOTE]
> カスタム関数ランタイムは、作業ウィンドウで使用されるランタイムの dialog オブジェクトとは少し異なるダイアログオブジェクトを使用します。 これらはどちらも "Dialog API" と呼ばれています`Officeruntime.Dialog`が、カスタム関数ランタイムでユーザーを認証するために使用します。

の`OfficeRuntime.Dialog`使用方法については、「[カスタム関数ランタイム](https://docs.microsoft.com/en-us/office/dev/add-ins/excel/custom-functions-runtime?view=office-js#displaying-a-dialog-box)」を参照してください。

全体として認証プロセス全体を構想する場合は、アドインの作業ウィンドウと UI 要素、およびアドインのカスタム関数部分を、を通じて`AsyncStorage`相互に通信できる個別のエンティティと考えることをお勧めします。

次の図は、この基本的なプロセスの概要を示しています。 点線では、個別のアクションを実行する一方で、カスタム関数とアドインの作業ウィンドウは、どちらもアドインの一部であることに注意してください。

1. Excel ブックのセルからカスタム関数呼び出しを発行します。
2. カスタム関数を使用`Officeruntime.Dialog`して、ユーザーの資格情報を web サイトに渡します。
3. その後、この web サイトは、カスタム関数へのアクセストークンを返します。
4. 次に、カスタム関数は、 `AsyncStorage`このアクセストークンをに設定します。
5. アドインの作業ウィンドウで、から`AsyncStorage`トークンにアクセスできます。

![カスタム関数、officeruntime、および共同作業ウィンドウの図](../images/Authdiagram.png "認証の図。")

## <a name="general-guidance"></a>一般的なガイダンス

Office アドインは web ベースであり、任意の web 認証方法を使用できます。 カスタム関数を使用して独自の認証を実装するために従う必要のある特定のパターンやメソッドはありません。 [外部サービスによる承認については、この記事](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/auth-external-add-ins?view=office-js)から始まるさまざまな認証パターンに関するドキュメントを参照してください。  

カスタム関数を開発する際に、次の場所を使用してデータを保存しないようにします。  

- `localStorage`: カスタム関数にはグローバル`window`オブジェクトへのアクセス権がないため、に`localStorage`格納されているデータにアクセスできません。
- `Office.context.document.settings`: この場所はセキュリティで保護されていないため、アドインを使用するすべてのユーザーが情報を抽出できます。

## <a name="see-also"></a>関連項目

* [カスタム関数のメタデータ](custom-functions-json.md)
* [Excel カスタム関数のランタイム](custom-functions-runtime.md)
* [カスタム関数のベスト プラクティス](custom-functions-best-practices.md)
* [チュートリアル: Excel でカスタム関数を作成します。](excel-tutorial-custom-functions.md)
