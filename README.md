
# Téléchargeur de pièces jointes d'email via Microsoft Graph

Ce script Python interagit avec l'API Microsoft Graph pour récupérer les emails et télécharger leurs pièces jointes. Le script filtre les emails reçus aujourd'hui et vérifie s'ils proviennent de certains expéditeurs. Si des pièces jointes sont trouvées, elles sont téléchargées et enregistrées dans un dossier nommé avec la date du jour sous le répertoire `attachments`.

## Prérequis

Pour exécuter ce script, vous devez avoir les éléments suivants :

- Python 3.6 ou supérieur
- Accès à l'API Microsoft Graph via une application Azure AD (pour l'authentification)
- Un fichier `.env` pour stocker vos identifiants

### Bibliothèques Python requises

Vous pouvez installer les bibliothèques nécessaires avec `pip` :

```bash
pip install msal requests python-dotenv pytz
```

## Configuration

### Créer une application Microsoft Azure AD :

1. Enregistrez une application dans le portail Azure.
2. Ajoutez les autorisations d'API pour Microsoft Graph, spécifiquement `Mail.Read` et `Mail.ReadWrite` sous **Autorisations déléguées** ou **Autorisations d'application**.
3. Obtenez le `Client ID`, le `Tenant ID`, et le `Client Secret` pour votre application.

### Créer un fichier .env

Le fichier `.env` stocke des informations sensibles telles que les identifiants de votre client et la boîte aux lettres que vous souhaitez accéder. Voici un exemple de la structure de votre fichier `.env` :

```
CLIENT_ID=your-client-id
CLIENT_SECRET=your-client-secret
TENANT_ID=your-tenant-id
SENDERS=sender1@example.com,sender2@example.com  # Liste séparée par des virgules des adresses email des expéditeurs
MAILBOX=your-mailbox@example.com  # La boîte aux lettres dont vous voulez récupérer les emails
```

## Fonctionnement

Le script s'authentifie d'abord avec l'API Microsoft Graph en utilisant le flux d'identification par identifiants d'application avec MSAL (Microsoft Authentication Library).
Il récupère les emails reçus aujourd'hui à l'aide du filtre `receivedDateTime`.
Les emails sont vérifiés pour les pièces jointes. Si des pièces jointes sont trouvées, elles sont téléchargées et enregistrées dans un dossier nommé avec la date du jour.
Le dossier est créé sous le répertoire `attachments`.

### Structure des dossiers

Les pièces jointes sont enregistrées dans la structure suivante :

```
/attachments
    /YYYY-MM-DD
        attachment1.pdf
        attachment2.jpg
```

Où `YYYY-MM-DD` correspond à la date actuelle lorsque le script est exécuté.

## Exécution du script

Pour exécuter le script, utilisez la commande suivante :

```bash
python run.py
```

Le script va :

1. S'authentifier avec l'API Microsoft Graph à l'aide des identifiants dans le fichier `.env`.
2. Récupérer les emails reçus aujourd'hui.
3. Filtrer les emails par les adresses des expéditeurs spécifiées dans le fichier `.env`.
4. Télécharger les pièces jointes de ces emails et les enregistrer dans un dossier nommé avec la date du jour.

## Dépannage

- Assurez-vous que le `Client ID`, le `Client Secret`, le `Tenant ID`, et les autres variables d'environnement sont correctement définis dans le fichier `.env`.
- Si vous rencontrez des problèmes d'authentification avec l'API, vérifiez que les autorisations requises sont accordées à votre application Azure.
- Assurez-vous que votre boîte aux lettres est accessible et que vous utilisez la bonne adresse email.

## Licence

Ce projet est sous licence MIT - voir le fichier LICENSE pour plus de détails.
