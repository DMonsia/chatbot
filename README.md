# API ChatBot
**API ChatBot** est une IA qui permet de lire, traiter et modifier des fichiers Excel via l'injection de macros VBA générées par OpenAi.


## Prérequis
- [x] Assurez-vous que `Docker` est installé sur votre machine. Suivez [ceci](https://github.com/DMonsia/hadoop-cluster/blob/main/docker-installation.md) sinon.
- [x] Assurez-vous que `make` est installé sur votre machine. Suivez [ceci](https://www.makeuseof.com/how-to-fix-make-command-not-found-error-ubuntu/) sinon.


## Installation
Pour installer l'API ChatBot, suivez les étapes suivantes :

1. Clonez le référentiel sur votre machine locale

Ouvrez ensuite le dossier du projet et créez un répertoire data dans le dossier `api`. Le dossier `./api/data` est utilisé pour stocker toutes les données générées par l'application, telle que l'historique des requetes `history.csv` et autres.

2. Créer une image Docker de l'API

Exécutez la commande suivante pour créer l'image API.
```bash
make build
```

3. Démarrer le conteneur de l'API

Utilisez ensuite la commande suivante pour démarrer le conteneur API Docker:
```bash
make run DATA_FILE=${PWD}/chatbot/data
```

4. Une fois le serveur opérationnel, vous pouvez accéder au swagger de l'API à l'adresse `localhost:8000/docs` dans votre navigateur Web.

Vous devriez maintenant pouvoir utiliser l'API. Si vous avez des questions ou des problèmes, n'hésitez pas à nous contacter à mdougban@yellowsys.fr.
