#!/bin/bash
# Script to create a new Alembic migration

if [ -z "$1" ]; then
    echo "Usage: ./scripts/create_migration.sh <migration_message>"
    echo "Example: ./scripts/create_migration.sh 'add users table'"
    exit 1
fi

echo "Creating migration: $1"
alembic revision --autogenerate -m "$1"
