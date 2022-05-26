---
title: Office アドインのTeams マニフェスト (プレビュー)
description: プレビュー JSON マニフェストの概要を確認します。
ms.date: 05/24/2022
ms.localizationpriority: high
ms.openlocfilehash: f5c529a982956922ae3a76de6d09e15e710f1f03
ms.sourcegitcommit: d06a37cd52f7389435bbbb3da3a90815ca2dce4a
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/25/2022
ms.locfileid: "65672103"
---
# <a name="teams-manifest-for-office-add-ins-preview"></a>Office アドインのTeams マニフェスト (プレビュー)

Microsoft は、Microsoft 365開発者プラットフォームについて多くの改善を行っています。 これらの機能強化により、Office アドインを含む、Microsoft 365 のすべての種類の拡張機能の開発、展開、インストール、および管理の一貫性が向上します。これらの変更は、既存のアドインと互換性があります。 

現在取り組んでいる重要な改善点の 1 つは、現在の JSON 形式のTeams マニフェストに基づいて、同じマニフェスト形式とスキーマを使用して、すべての Microsoft 365 拡張機能に対して 1 つの分布単位を作成する機能です。

これらの目標に向けた重要な第一歩を踏み出しました。これにより、Teams JSON マニフェストのバージョンを使用して、Windows でのみ実行される Outlook アドインを作成できます。

> [!NOTE]
> 新しいマニフェストはプレビューに使用でき、フィードバックに基づいて変更される可能性があります。 経験豊富なアドイン開発者には、それを試してみることをお勧めします。 プレビュー マニフェストは、運用環境のアドインでは使用しないでください。 

早期プレビュー期間中は、次の制限事項が適用されます。

- Teams マニフェストのプレビュー バージョンでは、Outlook アドインのみがサポートされ、Windows のサブスクリプション Office でのみサポートされます。 Excel、PowerPoint、Word へのサポートの拡張に取り組んでいます。
- アドインと Teams アプリ (Teams 個人用タブ、その他のMicrosoft 365拡張機能の種類など) を組み合わせてサイドロードすることは、まだできません。 今後数か月間、これらのシナリオをサポートするためにプレビューを拡張し続け、マニフェストをプレビュー形式に更新するための追加ツールを提供します。

> [!TIP]
> プレビュー Teams マニフェストの使用を開始する準備はできましたか? 「[Teams マニフェスト (プレビュー) を使用してOutlook アドインをビルドする](../quickstarts/outlook-quickstart-json-manifest.md)」から始めます。

## <a name="overview-of-the-json-manifest"></a>JSON マニフェストの概要

### <a name="schemas-and-general-points"></a>スキーマと一般的なポイント

[プレビュー JSON マニフェスト](/microsoftteams/platform/resources/dev-preview/developer-preview-intro.md)のスキーマは 1 つだけですが、現在の XML マニフェストには合計 7 つの[スキーマ](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8)があります。  

### <a name="conceptual-mapping-of-the-preview-json-and-current-xml-manifests"></a>プレビュー JSON マニフェストと現在の XML マニフェストの概念マッピング

このセクションでは、現在の XML マニフェストに精通している閲覧者向けのプレビュー JSON マニフェストについて説明します。 留意すべき点： 

- JSON では、XML のように属性と要素の値が区別されません。 通常、XML 要素にマップされる JSON は、要素値と各属性の両方を子プロパティにします。 次の例では、いくつかの XML マークアップとその同等の JSON を示します。
  
  ```xml
  <MyThing color="blue">Some text</MyThing>
  ```

  ```json
  "myThing" : {
      "color": "blue",
      "text": "Some text"
  }
  ```
- 現在の XML マニフェストには、複数の名前を持つ要素に、同じ名前の単数形バージョンの子が含まれている場所が多数あります。 たとえば、カスタム メニューを構成するマークアップには、複数の **Item** 要素の子を含めることができる **Items** 要素が含まれています。 これらの複数の要素に相当する JSON は、その値として配列を持つプロパティです。 配列のメンバーは *匿名* オブジェクトであり、"item" または "item1"、"item2" などの名前のプロパティではありません。次に例を示します。

  ```json
  "items": [
      {
          -- markup for a menu item is here --
      },
      {
          -- markup for another menu item is here --
      }
  ]
  ```

#### <a name="top-level-structure"></a>最上位レベルの構造

プレビュー JSON マニフェストのルート レベルは、現在の XML マニフェストの **OfficeApp** 要素とほぼ同じで、匿名オブジェクトです。 

**OfficeApp** の子は、一般的に 2 つの概念的なカテゴリに分類されます。 **VersionOverrides** 要素は 1 つのカテゴリです。 もう 1 つの子は、 **OfficeApp** の他のすべての子で構成され、これらはまとめて基本マニフェストと呼ばれます。 そのため、プレビュー JSON マニフェストにも同様の除算があります。 最上位レベルの "extension" プロパティがあり、その目的に大まかに対応し、子プロパティは **VersionOverrides** 要素に対応します。 プレビュー JSON マニフェストには、XML マニフェストの基本マニフェストと同じ目的でまとめて機能する 10 を超えるその他の最上位プロパティもあります。 これらの他のプロパティは、JSON マニフェストの基本マニフェストとまとめて考えることができます。 

> [!NOTE]
> 1 つのマニフェストでアドインを他の Microsoft 365 拡張機能の種類と組み合わせることができるようになると、基本マニフェストの概念に収まらない他の最上位レベルのプロパティが存在します。 通常、"configurableTabs"、"bots"、"connectors" など、あらゆる種類の Microsoft 365 拡張機能の種類に対して最上位のプロパティが存在します。 例については、[Teams マニフェストのドキュメント](/microsoftteams/platform/resources/schema/manifest-schema)を参照してください。 この構造により、"extension" プロパティは、Microsoft 365 拡張機能の 1 つの種類として Office アドインを表すことがわかります。

#### <a name="base-manifest"></a>基本マニフェスト

基本マニフェスト プロパティは、*どんな* 種類の Microsoft 365 拡張機能も持つことが想定されるアドインの特性を指定します。 これには、Office アドインだけでなく、Teams タブとメッセージ拡張機能も含まれます。これらの特性には、パブリック名と一意の ID が含まれます。 次の表は、プレビュー JSON マニフェストの重要な最上位プロパティと現在のマニフェストの XML 要素のマッピングを示しています。このマッピング原則がマークアップの *目的* です。

|JSON プロパティ|用途|XML 要素|コメント|
|:-----|:-----|:-----|:-----|
|"$schema"| マニフェスト スキーマを識別します。 | **OfficeApp** と **VersionOverrides** の属性 | |
|"id"| アドインの GUID。 | **Id**| |
|"version"| アドインのバージョンです。 | **バージョン** | |
|"manifestVersion"| マニフェスト スキーマのバージョンです。 |  **OfficeApp** の属性 | |
|"name"| アドインのパブリック名。 | **DisplayName** | |
|"description"| アドインの公開用の説明。  | **説明** | |
|"accentColor"||| このプロパティは、現在の XML マニフェストに同等のものはなく、JSON マニフェストのプレビューでは使用されません。 ただし、存在する必要があります。 |
|"developer"| アドインの開発者を識別します。 | **ProviderName** | |
|"localizationInfo"| 既定のロケールとその他のサポートされているロケールを構成します。 | **DefaultLocale** と **Override** | |
|"webApplicationInfo"| Azure Active Directory で既知のアドインの Web アプリを識別します。 | **WebApplicationInfo** | 現在の XML マニフェストでは、 **WebApplicationInfo** 要素は基本マニフェストではなく **VersionOverrides** 内にあります。 |
|"authorization"| アドインに必要な Microsoft Graphアクセス許可を識別します。 | **WebApplicationInfo** | 現在の XML マニフェストでは、 **WebApplicationInfo** 要素は基本マニフェストではなく **VersionOverrides** 内にあります。 |

**ホスト**、**要件**、**および ExtendedOverrides** 要素は、現在の XML マニフェストの基本マニフェストの一部です。 ただし、これらの要素に関連する概念と目的は、プレビュー JSON マニフェストの "extension" プロパティ内で構成されます。 

#### <a name="extension-property"></a>"extension" プロパティ

プレビュー JSON マニフェストの "extension" プロパティは、主に、他の種類の Microsoft 365 拡張機能には関連しないアドインの特性を表します。 たとえば、アドインが拡張するOffice アプリケーション (Excel、PowerPoint、Word、Outlook など) は、Office アプリケーション リボンのカスタマイズと同様に、"extension" プロパティ内で指定されます。 "extension" プロパティの構成目的は、現在の XML マニフェストの **VersionOverrides** 要素の構成と密接に一致します。

> [!NOTE]
> 現在の XML マニフェストの **VersionOverrides** セクションには、多くの文字列リソースに対する "二重ジャンプ" システムがあります。 URL を含む文字列が指定され、**VersionOverrides** の **Resources** 子に ID が割り当てられます。 文字列を必要とする要素には、**Resources** 要素内の文字列の ID と一致する `resid` 属性があります。 プレビュー JSON マニフェストの "extension" プロパティは、文字列をプロパティ値として直接定義することで、物事を簡略化します。 JSON マニフェストには、**Resources** 要素に相当するものはありません。

次の表は、プレビュー JSON マニフェストの "extension" プロパティの一部の高レベルの子プロパティと、現在のマニフェストの XML 要素のマッピングを示しています。 ドット表記は、子プロパティを参照するために使用されます。

|JSON プロパティ|用途|XML 要素|コメント|
|:-----|:-----|:-----|:-----|
| "requirements.capabilities" | アドインをインストール可能にする必要がある要件セットを識別します。 | **要件** と **セット** | |
| "requirements.scopes" | アドインをインストールできる Office アプリケーションを識別します。 | **Hosts** |  |
| "ribbons" | アドインがカスタマイズするリボン。 | **ホスト**、 **ExtensionPoints**、およびさまざまな **\*FormFactor** 要素 | "ribbons" プロパティは、これら 3 つの要素の目的をマージする匿名オブジェクトの配列です。 [「リボン」の表](#ribbons-table)を参照してください。|
| "alternatives" | 同等の COM アドイン、XLL、またはその両方との下位互換性を指定します。 | **EquivalentAddins** | 背景情報については、[「EquivalentAddins - 参照」](/javascript/api/manifest/equivalentaddins#see-also)を参照してください。 |
| "runtimes"  | カスタム のリボン ボタンから直接実行されるカスタム関数や関数など、さまざまな種類の "UI のない" アドインを構成します。 | **Runtimes**。 **FunctionFile**、および **ExtensionPoint** (CustomFunctions 型) |  |
| "autoRunEvents" | 指定したイベントのイベント ハンドラーを構成します。 | **Event** と **ExtensionPoint** (イベントの種類) |  |

##### <a name="ribbons-table"></a>"ribbons" テーブル

次の表は、"ribbons" 配列内の匿名子オブジェクトの子プロパティを、現在のマニフェストの XML 要素にマップします。 

|JSON プロパティ|用途|XML 要素|コメント|
|:-----|:-----|:-----|:-----|
| "contexts" | アドインがカスタマイズするコマンド サーフェスを指定します。 | **PrimaryCommandSurface** や **MessageReadCommandSurface** など、さまざまな **\*CommandSurface** 要素 |  |
| "tabs" | カスタム リボン タブを構成します。 | **CustomTab** | "タブ" の子孫プロパティの名前と階層は、**CustomTab** の子孫と密接に一致します。  |

## <a name="sample-preview-json-manifest"></a>サンプル プレビュー JSON マニフェスト

アドインのプレビュー JSON マニフェストの例を次に示します。

```json
{
  "$schema": "https://raw.githubusercontent.com/OfficeDev/microsoft-teams-app-schema/op/extensions/MicrosoftTeams.schema.json",
  "id": "00000000-0000-0000-0000-000000000000",
  "version": "1.0.0",
  "manifestVersion": "devPreview",
  "name": {
    "short": "Name of your app (<=30 chars)",
    "full": "Full name of app, if longer than 30 characters (<=100 chars)"
  },
  "description": {
    "short": "Short description of your app (<= 80 chars)",
    "full": "Full description of your app (<= 4000 chars)"
  },
  "icons": {
    "outline": "outline.png",
    "color": "color.png"
  },
  "accentColor": "#230201",
  "developer": {
    "name": "Contoso",
    "websiteUrl": "https://www.contoso.com",
    "privacyUrl": "https://www.contoso.com/privacy",
    "termsOfUseUrl": "https://www.contoso.com/servicesagreement"
  },
  "localizationInfo": {
    "defaultLanguageTag": "en-us",
    "additionalLanguages": [
      {
        "languageTag": "es-es",
        "file": "es-es.json"
      }
    ]
  },
  "webApplicationInfo": {
    "id": "00000000-0000-0000-0000-000000000000",
    "resource": "api://www.contoso.com/prodapp"
  },
  "authorization": {
    "permissions": {
      "resourceSpecific": [
        {
          "name": "Mailbox.ReadWrite.User",
          "type": "Delegated"
        }
      ]
    }
  },
  "extensions": [
    {
      "requirements": {
        "scopes": [ "mail" ],
        "capabilities": [
          {
            "name": "Mailbox", "minVersion": "1.1"
          }
        ]
      },
      "runtimes": [
        {
          "requirements": {
            "capabilities": [
              {
                "name": "MailBox", "minVersion": "1.10"
              }
            ]
          },
          "id": "eventsRuntime",
          "type": "general",
          "code": {
            "page": "https://contoso.com/events.html",
            "script": "https://contoso.com/events.js"
          },
          "lifetime": "short",
          "actions": [
            {
              "id": "onMessageSending",
              "type": "executeFunction"
            },
            {
              "id": "onNewMessageComposeCreated",
              "type": "executeFunction"
            }
          ]
        },
        {
          "requirements": {
            "capabilities": [
              {
                "name": "MailBox", "minVersion": "1.1"
              }
            ]
          },
          "id": "commandsRuntime",
          "type": "general",
          "code": {
            "page": "https://contoso.com/commands.html",
            "script": "https://contoso.com/commands.js"
          },
          "lifetime": "short",
          "actions": [
            {
              "id": "action1",
              "type": "executeFunction"
            },
            {
              "id": "action2",
              "type": "executeFunction"
            },
            {
              "id": "action3",
              "type": "executeFunction"
            }
          ]
        }
      ],
      "ribbons": [
        {
          "contexts": [
            "mailCompose"
          ],
          "tabs": [
            {
              "builtInTabId": "TabDefault",
              "groups": [
                {
                  "id": "dashboard",
                  "label": "Controls",
                  "controls": [
                    {
                      "id": "control1",
                      "type": "button",
                      "label": "Action 1",
                      "icons": [
                        {
                          "size": 16,
                          "file": "test_16.png"
                        },
                        {
                          "size": 32,
                          "file": "test_32.png"
                        },
                        {
                          "size": 80,
                          "file": "test_80.png"
                        }
                      ],
                      "supertip": {
                        "title": "Action 1 Title",
                        "description": "Action 1 Description"
                      },
                      "actionId": "action1"
                    },
                    {
                      "id": "menu1",
                      "type": "menu",
                      "label": "My Menu",
                      "icons": [
                        {
                          "size": 16,
                          "file": "test_16.png"
                        },
                        {
                          "size": 32,
                          "file": "test_32.png"
                        },
                        {
                          "size": 80,
                          "file": "test_80.png"
                        }
                      ],
                      "supertip": {
                        "title": "My Menu",
                        "description": "Menu with 2 actions"
                      },
                      "items": [
                        {
                          "id": "menuItem1",
                          "type": "menuItem",
                          "label": "Action 2",
                          "supertip": {
                            "title": "Action 2 Title",
                            "description": "Action 2 Description"
                          },
                          "actionId": "action2"
                        },
                        {
                          "id": "menuItem2",
                          "type": "menuItem",
                          "label": "Action 3",
                          "icons": [
                            {
                              "size": 16,
                              "file": "test_16.png"
                            },
                            {
                              "size": 32,
                              "file": "test_32.png"
                            },
                            {
                              "size": 80,
                              "file": "test_80.png"
                            }
                          ],
                          "supertip": {
                            "title": "Action 3 Title",
                            "description": "Action 3 Description"
                          },
                          "actionId": "action3"
                        }
                      ]
                    }
                  ]
                }
              ]
            }
          ]
        },
        {
          "contexts": [ "mailRead" ],
          "tabs": [
            {
              "builtInTabId": "TabDefault",
              "groups": [
                {
                  "id": "dashboard",
                  "label": "Controls",
                  "controls": [
                    {
                      "id": "control1",
                      "type": "button",
                      "label": "Action 1",
                      "icons": [
                        {
                          "size": 16,
                          "file": "test_16.png"
                        },
                        {
                          "size": 32,
                          "file": "test_32.png"
                        },
                        {
                          "size": 80,
                          "file": "test_80.png"
                        }
                      ],
                      "supertip": {
                        "title": "Action 1 Title",
                        "description": "Action 1 Description"
                      },
                      "actionId": "action1"
                    }
                  ]
                }
              ]
            }
          ]
        }
      ],
      "autoRunEvents": [
        {
          "requirements": {
            "capabilities": [
              {
                "name": "MailBox", "minVersion": "1.10"
              }
            ]
          },
          "events": [
            {
              "type": "newMessageComposeCreated",
              "actionId": "onNewMessageComposeCreated"
            },
            {
              "type": "messageSending",
              "actionId": "onMessageSending",
              "options": {
                "sendMode": "promptUser"
              }
            }
          ]
        }
      ],
      "alternates": [
        {
          "requirements": {
            "scopes": [ "mail" ]
          },
          "prefer": {
            "comAddin": {
              "progId": "ContosoExtension"
            }
          },
          "hide": {
            "storeOfficeAddin": {
              "officeAddinId": "00000000-0000-0000-0000-000000000000",
              "assetId": "WA000000000"
            }
          }
        }
      ]
    }
  ]
}
```

## <a name="next-steps"></a>次の手順

- [Teams マニフェスト (プレビュー) を使用して、Outlook アドインをビルド](../quickstarts/outlook-quickstart-json-manifest.md)します。