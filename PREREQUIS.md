# Add-in Office 365 Prisme.ai — pré-requis (synthèse)

_Rédigé le 2026-10-12 à partir de la doc officielle Microsoft (Office Add-ins,
Microsoft 365 admin) et de la doc publique « Claude for Microsoft 365 »
(Anthropic), qui sert de référence d'architecture. Sources en bas de page._

## 1. Ce qu'est un add-in Office

Deux composants, rien d'autre :

| Composant | Rôle | Contrainte |
|---|---|---|
| **Manifest** | Déclare l'add-in à Office : id, hôtes (Word/Excel/PowerPoint/Outlook), bouton de ruban, URL du volet, permissions | Déployé par un admin M365 (ou sideloadé en dev) |
| **Web app** | La page chargée dans le volet (task pane) + `office.js` | **Hébergée en HTTPS** (obligatoire pour Office sur le web et le store) |

Le volet est une webview : **WebView2** (Windows), **WKWebView/Safari** (Mac, iOS),
ou une **iframe dans le navigateur** (Office sur le web).

## 2. Décisions d'architecture (alignées sur ce que fait Claude)

| Sujet | Choix | Pourquoi |
|---|---|---|
| Type de manifest | **XML « add-in only »** | Le manifest unifié (JSON) n'est pas supporté sur Outlook Mac, mobile, Office perpétuel, et exige Word/Excel/PPT Windows ≥ 2501 et Mac ≥ 16.103. Claude livre aussi du XML. |
| Nombre de manifests | **2** : `manifest.xml` (Word + Excel + PowerPoint) et `manifest-outlook.xml` | Outlook exige un manifest séparé (type `MailApp`). Même découpage chez Claude. |
| UI du chat | iframe `/embed/chat` de l'instance, même contrat `postMessage` que l'extension Chrome | Zéro code de chat dupliqué ; le token reste dans l'iframe. |
| Config d'instance | Paramètres dans l'URL du volet, écrits dans le manifest (`?instance=…&api=…&agent=…`) | Équivalent de `chrome.storage.managed`. Claude fait pareil (`pivot.claude.ai/?gateway_url=…`). **Jamais de secret dans la query string.** |
| Accès au contenu | `office.js` uniquement (document / mail ouvert) | Pas de Microsoft Graph pour lire du contenu. Une app Entra n'est requise que pour le SSO. |

## 3. Pré-requis

### 3.1 Hébergement du volet (à fournir par Prisme.ai ou le client)

- [ ] Un hôte **HTTPS** avec certificat valide servant `taskpane.html`, `taskpane.js`, `icons/` (statique, pas de backend).
- [ ] Pas de `X-Frame-Options` / `frame-ancestors` restrictif sur cet hôte (Office sur le web charge le volet en iframe).
- [ ] Icônes servies **avec cache** (pas de `Cache-Control: no-cache/no-store`).
- [ ] Nom d'hôte **stable** : l'identité de l'add-in suit l'URL du manifest.

### 3.2 Instance Prisme.ai

- [ ] **Platform URL**, **API URL**, **Agent ID** (Settings → Access Tokens).
- [ ] Mode d'auth : accès **anonyme** autorisé pour l'agent (+ `orgSlug` si l'org le restreint), ou **SSO Entra** (voir point ouvert 1 et le README).
- [ ] Framing de `/embed/chat` : ouvert par défaut (`frame-ancestors *`). Si `EMBED_ALLOWED_ORIGINS` est défini sur le service console : y ajouter l'origine de l'hôte du volet **et** les origines d'Office sur le web (`frame-ancestors` vérifie toute la chaîne d'ancêtres ; les jokers CSP type `https://*.officeapps.live.com` sont acceptés). _Liste exacte des origines Office à valider sur une instance réelle._
- [ ] Pour les **actions** (écrire dans le doc, brouillon de réponse) : un workspace d'outils `office-mcp` côté instance (pendant de `browser-mcp`), dans le repo `prismeai-workspaces` — à importer sur l'instance.

### 3.3 Tenant Microsoft 365 (déploiement organisation)

- [ ] **Rôle admin** : Global Admin, ou Exchange Admin + Application Administrator. ⚠ Les rôles activés via **PIM** ne sont pas reconnus par « Integrated apps » (bug Microsoft connu) → rôle actif permanent.
- [ ] **Licences utilisateurs** : M365 Business Basic/Standard/Premium, Office 365 E1/E3/E5/F3, M365 E3/E5/F3 (ou équivalents Education / Government).
- [ ] **Exchange Online** avec boîte active pour chaque utilisateur (les manifests y sont stockés). Les SKU sans Exchange Online (M365 Apps for Business/Enterprise seuls) ne sont pas supportés.
- [ ] Annuaire dans **Entra ID** (ou fédéré). Pas d'annuaire on-prem seul.
- [ ] Groupes de **premier niveau** uniquement (pas de groupes imbriqués, pas de groupes de sécurité non mail-enabled).
- [ ] Prévoir **jusqu'à 24 h** de propagation après déploiement.

### 3.4 Postes utilisateurs

| Client | Minimum (déploiement centralisé) |
|---|---|
| Word / Excel / PowerPoint Windows | Abonnement M365, version ≥ 1704 |
| Word / Excel / PowerPoint Mac | ≥ 15.34 |
| Office sur le web | Navigateur moderne (Edge, Chrome, Firefox, Safari) |
| Outlook Windows (classique et nouveau), Mac, web | Abonnement M365, boîte Exchange Online |

Non supporté / à exclure du périmètre : Office 2016/2019 perpétuel (MSI),
comptes POP3/IMAP, boîtes Exchange on-prem (hors déploiement Exchange Server
2016+ dédié), mobile en v0.

### 3.5 Réseau (allowlist proxy / pare-feu)

- [ ] `appsforoffice.microsoft.com` — runtime `office.js` (obligatoire pour tout add-in).
- [ ] L'hôte du volet (3.1).
- [ ] La **Platform URL** et l'**API URL** de l'instance Prisme.ai (HTTPS + WebSocket).

### 3.6 Poste de dev

- [ ] Node.js ≥ 18, un compte M365 de test (ou tenant dev).
- [ ] Certificat localhost approuvé (`npm run certs`) — Office refuse le HTTP et les certificats non approuvés.

## 4. Canaux de distribution

| Canal | Usage | Remarques |
|---|---|---|
| **Sideloading** | Dev / test | Pas pour la prod. |
| **Integrated apps** (admin center → Settings → Integrated apps → Upload custom apps → Office Add-in) | **Prod en entreprise** — canal cible | Upload du fichier manifest, assignation utilisateurs/groupes. Un upload par manifest (donc 2). Indisponible dans les clouds souverains / gouvernementaux → « Centralized Deployment » (Settings → Add-ins). |
| **Microsoft Marketplace** (ex-AppSource) | Diffusion publique | Compte Partner Center + certification Microsoft ; l'add-in doit fonctionner sur toutes les plateformes déclarées. Peu compatible avec une URL d'instance par client. |

Toute modification du manifest ⇒ **incrémenter `<Version>`** ; un changement de
permissions redemande le consentement de l'admin.

## 5. Points ouverts

1. **SSO** : implémenté par échange de token (`auth: 'sso'`), sans changement de code plateforme — id_token Entra obtenu via MSAL Nested App Authentication, échangé sur `POST /v2/login/token-exchange`, puis passé à l'embed en `auth: 'token'`. Contraintes : fournisseur SSO Prisme en OIDC sur Entra ID, « Allow Token Exchange » activé, **même app Entra** que le SSO web (le `sub` Entra est propre à chaque app et c'est la clé de rapprochement des comptes), origine du volet dans `CORS_ADDITIONAL_ALLOWED_ORIGINS` de l'api-gateway. **Non testé dans Office.** Les autres IdP restent en anonyme (piste : client PKCE via `OIDC_EXTERNAL_CLIENTS` + dialogue Office).
2. **Workspace `office-mcp`** : créé dans `prismeai-workspaces/workspaces/office-mcp` (calqué sur `browser-mcp`), pas encore déployé ni testé sur une instance. Voir `README.md` → Agent tools.
3. **Stockage partitionné** : l'iframe du chat est tierce dans le volet (ITP Safari / partitionnement Chromium) → la session ne survit pas forcément à la fermeture du volet.
4. **Recherche dans toute la boîte mail / agenda** : nécessiterait Microsoft Graph, donc une app Entra et un admin consent (c'est ce que fait Claude for Outlook : `Mail.ReadWrite`, `Calendars.Read`, `User.Read`, `offline_access`). Hors v0.
5. **Injection de prompt** : le contenu d'un document ou d'un mail est une donnée non fiable ; les outils d'écriture doivent rester soumis aux `tool_permissions` de l'agent.

## Sources

Microsoft
- [Office Add-ins platform overview](https://learn.microsoft.com/en-us/office/dev/add-ins/overview/office-add-ins)
- [Office Add-ins manifest](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/add-in-manifests)
- [Unified manifest — client and platform support](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/unified-manifest-overview)
- [Requirements for running Office Add-ins](https://learn.microsoft.com/en-us/office/dev/add-ins/concepts/requirements-for-running-office-add-ins)
- [Browsers and webview controls used by Office Add-ins](https://learn.microsoft.com/en-us/office/dev/add-ins/concepts/browsers-used-by-office-web-add-ins)
- [Deploy and publish Office Add-ins](https://learn.microsoft.com/en-us/office/dev/add-ins/publish/publish)
- [Integrated apps portal](https://learn.microsoft.com/en-us/microsoft-365/admin/manage/test-and-deploy-microsoft-365-apps)
- [Requirements for Centralized Deployment](https://learn.microsoft.com/en-us/microsoft-365/admin/manage/centralized-deployment-of-add-ins)

Anthropic (référence d'architecture)
- [Use Claude for Excel](https://claude.com/docs/office-agents/excel)
- [Use Claude for Outlook](https://claude.com/docs/office-agents/outlook)
- [Use Claude for M365 with third-party platforms](https://claude.com/docs/office-agents/third-party-platforms)
- [Work across M365 apps](https://claude.com/docs/office-agents/work-across-apps)
- [Deploy Claude for Microsoft 365 in Claude for Government](https://claude.com/docs/government/deploy-microsoft-365/configure)
