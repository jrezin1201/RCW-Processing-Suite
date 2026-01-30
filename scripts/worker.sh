#!/bin/bash
# Script to start RQ worker

echo "Starting RQ worker..."
rq worker default --with-scheduler
