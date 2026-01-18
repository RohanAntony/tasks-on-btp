export type TaskStatus = 'open' | 'in_progress' | 'review' | 'completed'

export interface Task {
  ID: string
  title: string
  description: string
  dueDate: string
  status: TaskStatus
}
