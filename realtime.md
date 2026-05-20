# Real-Time Processing — Annual Rollup System

## 🧠 Purpose

Defines how datasets are processed through the system in near real-time batch workflows.

---

## ⚡ Event Flow Architecture

```mermaid
sequenceDiagram
participant S as Source System
participant I as Ingestion Layer
participant A as Aggregation Engine
participant R as Rollup Processor
participant O as Output System

S->>I: Send Dataset
I->>A: Validate + Normalize
A->>R: Aggregate Metrics
R->>O: Generate Report
O-->>S: Return Output
```

---

## 🧩 Real-Time Behavior

- Batch-based near real-time processing
- Event-driven pipeline execution
- Incremental aggregation support
- Immediate report availability after processing

---

## 🎯 Design Goal

Ensure structured datasets are processed into accurate annual summaries with minimal delay.
