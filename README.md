# API ChatBot
**API ChatBot** est une IA qui permet de lire, traiter et modifier des fichiers Excel via l'injection de macros VBA générées par OpenAi.


## Prérequis
- [x] Assurez-vous que `Windows` est l'os de votre machine.
- [x] Assurez-vous que `Excel` est installé sur votre machine.

## Installation
Pour installer l'API ChatBot, suivez les étapes suivantes :

1. Clonez le référentiel sur votre machine locale

Ouvrez ensuite le dossier du projet et créez un répertoire data dans le dossier `chatbot`. Le dossier `./chatbot/data` est utilisé pour stocker toutes les données générées par l'application, telle que l'historique des requetes `history.csv` et autres.

2. Installer les dépendances de l'application
Toutes les dépendances Python nécessaires pour l'exécution de l'api se trouve dans le fichier `chatbot/requirements.txt`. Il est fortement conseillé de créer un environnement virtuel dédié au projet.

3. Démarrer le conteneur de l'API
Une fois dans le dossier chatbot entrer la commande:
```bash
uvicorn main:app
```

4. Une fois le serveur opérationnel, vous pouvez accéder au swagger de l'API à l'adresse `localhost:8000/docs` dans votre navigateur Web.

Vous devriez maintenant pouvoir utiliser l'API. Si vous avez des questions ou des problèmes, n'hésitez pas à nous contacter à mdougban@yellowsys.fr.
