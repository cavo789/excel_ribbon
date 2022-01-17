#! /bin/bash

docker run --rm -v "${PWD}"/:/data bosa/writing-doc concat -o README.MD
