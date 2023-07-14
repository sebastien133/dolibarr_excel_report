Installation :

Lancer en ligne de commande 'composer install' puis copier et renommer .env.example en .env

Mettre à jour les variables dans le fichier .env.

./script.sh devrait s'exécuter. Vérifier que vous avez donné les droits d'exécution au script. ("sudo chmod +x script.sh")

Envoi de mails :

L'envoi de mail se fait de manière automatique. Il est possible de conserver les fichiers sur le serveur en mettant la variable d'environnment ONLY_MAILS à false dans le fichier .env.