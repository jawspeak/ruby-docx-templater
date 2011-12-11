#!/usr/bin/env bash

bundler_version="1.0.21"

source "$HOME/.rvm/scripts/rvm"

function install_ruby_if_needed() {
  echo "Checking for $1..."
  if ! rvm list rubies | grep $1 > /dev/null; then
    rvm install $1
  fi
}
function switch_ruby() {
  install_ruby_if_needed $1 && rvm use $1
}

function install_bundler_if_needed() {
  echo "Checking for Bundler $bundler_version..."
  if ! gem list --installed bundler --version "$bundler_version" > /dev/null; then
    gem install bundler --version "$bundler_version" --source http://gems.rubyforge.org/
  fi
}

function update_gems_if_needed() {
  echo "Installing gems..."
  bundle check || bundle install
}

function run_specs() {
  # TMPDIR is a work-around for http://jira.codehaus.org/browse/JRUBY-4033
  env TMPDIR="tmp" bundle exec rake spec
}

function prepare_and_run() {
  switch_ruby $1 &&
  install_bundler_if_needed &&
  update_gems_if_needed &&
  run_specs
}

function tag_green_build() {
  tag_name="ci-ruby-docx-creator-master/latest"
  git tag -f -m "tagging green build" "$tag_name"
  git push -f origin "$tag_name"
}

prepare_and_run "ree-1.8.7-2011.03" &&
tag_green_build