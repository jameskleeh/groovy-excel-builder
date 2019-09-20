#!/bin/bash
set -e

./gradlew clean check

EXIT_STATUS=0

echo "Publishing archives for branch $TRAVIS_BRANCH"
if [[ -n $TRAVIS_TAG ]] || [[ $TRAVIS_BRANCH =~ ^master|[012]\..\.x$ && $TRAVIS_PULL_REQUEST == 'false' ]]; then

  if [[ -n $TRAVIS_TAG ]]; then
    echo "Pushing build to Bintray"

    ./gradlew bintrayUpload || EXIT_STATUS=$?

    if [[ $EXIT_STATUS == 0 ]]; then
        ./gradlew gitPublishPush || EXIT_STATUS=$?
    fi

  else
    echo "Publishing snapshot to OJO"

    ./gradlew artifactoryPublish || EXIT_STATUS=$?
  fi

fi

exit $EXIT_STATUS
