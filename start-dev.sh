#!/bin/bash

# Запуск бэкенда
cd backend
npm run dev &

# Запуск фронтенда
cd ../frontend
npm start &

# Ожидание завершения обоих процессов
wait
