#!/bin/sh

GIT=`which git`
REPO_DIR=/Users/florianzasada/Library/CloudStorage/GoogleDrive-florian.zasada@gmail.com/Meine Ablage/CapZa/capza
cd ${REPO_DIR}
${GIT} add --all .
${GIT} commit -m "run update"
${GIT} push git@bitbucket.org:FlorianZasa/capza.git master