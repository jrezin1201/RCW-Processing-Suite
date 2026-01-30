#!/bin/bash
# Script to run pending Alembic migrations

echo "Running database migrations..."
alembic upgrade head
echo "Migrations complete!"
