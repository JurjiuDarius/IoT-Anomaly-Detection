# IoT Anomaly Detection System

Real-time anomaly detection over streaming IoT sensor data — an end-to-end microservices pipeline built for my MSc Big Data course.

![Kafka](https://img.shields.io/badge/Apache%20Kafka-231F20?logo=apachekafka&logoColor=white)
![Hadoop](https://img.shields.io/badge/HDFS-66CCFF?logo=apachehadoop&logoColor=black)
![Python](https://img.shields.io/badge/Python-3776AB?logo=python&logoColor=white)
![scikit-learn](https://img.shields.io/badge/scikit--learn-F7931E?logo=scikitlearn&logoColor=white)
![Streamlit](https://img.shields.io/badge/Streamlit-FF4B4B?logo=streamlit&logoColor=white)
![Docker](https://img.shields.io/badge/Docker-2496ED?logo=docker&logoColor=white)

Temperature and vibration sensors stream into **Kafka**, an **Isolation Forest** detector flags anomalies in real time, raw data lands in **HDFS** for history, and a live **Streamlit** dashboard visualizes both the stream and the batch view. The whole stack comes up with a single `docker compose up`.

## Architecture

```
sensors → Kafka (sensors.raw) → detector ─┬─→ Kafka (sensors.anomalies) → dashboard (live)
                                          └─→ HDFS (historical) ─────────→ dashboard (batch)
```

1. **Data source** — `src/producers/main.py` simulates temperature & vibration sensors (with occasional injected anomalies).
2. **Ingestion** — Apache Kafka `sensors.raw` topic.
3. **Processing** — `src/detector/main.py` runs an Isolation Forest, publishes hits to `sensors.anomalies`, and batches raw data to HDFS.
4. **Storage** — HDFS for historical data.
5. **Visualization** — `src/dashboard/app.py`, a Streamlit dashboard for live monitoring and batch analysis.

## Quickstart

```bash
docker compose up --build -d          # Kafka + HDFS take a minute to settle
open http://localhost:8501            # the dashboard
docker compose down                   # tear down
```

In the dashboard: tick **Run Live Monitoring** for the real-time view, or open **Batch Analysis → Load Data from HDFS** for historical stats once some data has accumulated.

## Components

Producer · Isolation Forest detector · Streamlit dashboard · Kafka + Zookeeper · HDFS namenode + datanode — each a container in `docker-compose.yml`.

## Tech stack

Apache Kafka · Hadoop HDFS · Python · scikit-learn (Isolation Forest) · Streamlit · Docker Compose

---

*Developed as a team project for the MSc Big Data course.*
