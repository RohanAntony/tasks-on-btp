export type TaskStatus = 'open' | 'in_progress' | 'review' | 'completed'

export interface TaskHistory {
  ID: string
  field: string
  oldValue: string | null
  newValue: string | null
  createdAt: string
}

export interface Task {
  ID: string
  title: string
  description: string
  dueDate: string
  status: TaskStatus
  createdAt: string
}
