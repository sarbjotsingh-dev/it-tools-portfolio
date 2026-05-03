# Employee Quiz Platform

A two-view quiz system — an **Admin dashboard** to create and manage quizzes, and an **Agent view** for employees to take them. Results are stored in SharePoint lists.

## Features

- **Admin View** — create questions, set passing scores, view per-agent results and attempt history
- **Agent View** — clean quiz-taking interface with timer, progress indicator, and instant score on submit
- **Live Scoring** — answers evaluated on submit, score persisted to SharePoint
- **Pass/Fail Logic** — configurable passing threshold with visual feedback
- **Results History** — admins can view all attempts per agent with timestamps

## Tech Stack

- Vanilla HTML / CSS / JavaScript
- SharePoint REST API (lists for questions, answers, results)
- Fluent UI-inspired design

## Setup

1. Deploy `quiz-admin.html`, `quiz-agent.html`, and `quiz.js` to SharePoint Site Assets
2. Update `SITE_URL` in `quiz.js`
3. Run `create-quiz-lists.ps1` to scaffold the required SharePoint lists, or create them manually:
   - `QuizQuestions` — Id, Title, Options (JSON), CorrectAnswer, QuizId
   - `QuizResults` — AgentName, Score, Passed, AttemptDate, QuizId
