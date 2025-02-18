
# Téléchargeur de pièces jointes d'e-mails Microsoft Graph

Cette application Flask permet de s'authentifier via Microsoft OAuth2, de récupérer les e-mails d'une boîte aux lettres spécifiée et de télécharger les pièces jointes des e-mails provenant d'expéditeurs spécifiques. Les pièces jointes sont regroupées dans un fichier ZIP pour un téléchargement facile.

## Prérequis

Pour exécuter ce script, vous aurez besoin des éléments suivants :

- Un compte Microsoft Azure
- Un enregistrement d'application Azure avec les autorisations nécessaires
- Python 3.7 ou version ultérieure
- Pip pour installer les dépendances

## Configuration

### 1. Enregistrer votre application dans Azure AD

Avant d'exécuter le script, vous devez enregistrer votre application dans Azure Active Directory (Azure AD) et définir les autorisations API nécessaires :

- Allez sur le [portail Azure](https://portal.azure.com).
- Accédez à **Azure Active Directory** > **Enregistrements d'applications** > **Nouvel enregistrement**.
- Notez l'**ID de l'application (client)**, l'**ID du locataire (tenant)** et le **Secret du client**. Ces informations seront nécessaires pour configurer le script.

### 2. Installer les dépendances

Assurez-vous que `pip` est installé, puis installez les dépendances requises :

```bash
pip install -r requirements.txt
```

Cela installera :

- `Flask` pour le framework web.
- `requests` pour effectuer des requêtes HTTP.
- `msal` pour gérer l'authentification Microsoft OAuth.
- `python-dotenv` pour charger les variables d'environnement.
- `pytz` pour gérer les fuseaux horaires.

### 3. Créer un fichier `.env`

Créez un fichier `.env` dans le répertoire racine du projet et ajoutez les détails suivants :

```bash
CLIENT_ID=<votre-client-id>
CLIENT_SECRET=<votre-client-secret>
TENANT_ID=<votre-tenant-id>
MAILBOX=<votre-email-boîte-aux-lettres>
SENDERS=<emails-expéditeurs-séparés-par-virgules>
```

- `CLIENT_ID` : L'ID de l'application Azure.
- `CLIENT_SECRET` : Le secret client généré pour votre application Azure.
- `TENANT_ID` : L'ID du locataire Azure AD.
- `MAILBOX` : L'adresse e-mail de la boîte aux lettres à partir de laquelle vous souhaitez récupérer les e-mails.
- `SENDERS` : Une liste d'expéditeurs, séparée par des virgules, dont vous souhaitez télécharger les pièces jointes.

### 4. Exécuter l'application

Démarrez le serveur Flask en exécutant :

```bash
python app.py
```

Par défaut, l'application sera accessible à l'adresse `http://localhost:3000`.

### 5. Authentification OAuth2

Lorsque vous accédez à `http://localhost:3000`, vous serez redirigé vers la page de connexion OAuth2 de Microsoft. Après vous être connecté et avoir accepté les autorisations demandées, vous serez redirigé vers l'application. Celle-ci récupérera les e-mails de la boîte aux lettres spécifiée et téléchargera les pièces jointes des expéditeurs spécifiés.

### 6. Télécharger les pièces jointes

Après une authentification réussie, l'application traitera les e-mails reçus au cours de la journée en cours et téléchargera les pièces jointes des expéditeurs spécifiés. Les pièces jointes seront enregistrées dans un dossier avec la date du jour, puis regroupées dans un fichier ZIP, qui sera proposé en téléchargement.

## Points de terminaison

- `/` : Lance le flux d'authentification OAuth2.
- `/token` : Gère la redirection de Microsoft après l'authentification et récupère le jeton d'accès.
- `/download` : Récupère et télécharge les pièces jointes des e-mails provenant des expéditeurs spécifiés.

## Exemple de flux

1. Accédez à `http://localhost:3000/`.
2. Authentifiez-vous via Microsoft OAuth2.
3. Après l'authentification, l'application récupère les e-mails et les pièces jointes des expéditeurs spécifiés.
4. Les pièces jointes sont regroupées dans un fichier ZIP et proposées pour téléchargement.

## Dépannage

- **Échec de l'autorisation** : Assurez-vous que vous avez entré les bons identifiants (client ID, client secret, tenant ID et boîte aux lettres).
- **Aucune pièce jointe trouvée** : Vérifiez que les expéditeurs spécifiés ont des e-mails avec des pièces jointes dans la période que vous interrogez.
- **Jeton d'accès invalide** : Ré-authentifiez-vous pour garantir l'utilisation d'un jeton valide.

## Licence

Ce projet est sous licence MIT.
