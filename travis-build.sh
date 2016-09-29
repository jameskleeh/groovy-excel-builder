#!/bin/bash
set -e
./gradlew clean test 
EXIT_STATUS=0
echo "Publishing archives for branch $TRAVIS_BRANCH"
if [[ $TRAVIS_PULL_REQUEST == 'false' ]]; then

  echo "Publishing archives"

  if [[ -n $TRAVIS_TAG ]]; then
      ./gradlew bintrayUpload || EXIT_STATUS=$?
  fi

fi

exit $EXIT_STATUS
