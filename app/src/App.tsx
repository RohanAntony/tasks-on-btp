import '@ui5/webcomponents-react/dist/Assets.js'
import { useEffect, useState } from 'react'
import { ThemeProvider } from '@ui5/webcomponents-react/ThemeProvider'
import { ShellBar } from '@ui5/webcomponents-react/ShellBar'
import { ShellBarItem } from '@ui5/webcomponents-react/ShellBarItem'
import { Table } from '@ui5/webcomponents-react/Table'
import { TableHeaderRow } from '@ui5/webcomponents-react/TableHeaderRow'
import { TableHeaderCell } from '@ui5/webcomponents-react/TableHeaderCell'
import { TableRow } from '@ui5/webcomponents-react/TableRow'
import { TableCell } from '@ui5/webcomponents-react/TableCell'
import { Tag } from '@ui5/webcomponents-react/Tag'
import { BusyIndicator } from '@ui5/webcomponents-react/BusyIndicator'
import { MessageStrip } from '@ui5/webcomponents-react/MessageStrip'
import { FlexBox } from '@ui5/webcomponents-react/FlexBox'
import { Title } from '@ui5/webcomponents-react/Title'
import addIcon from '@ui5/webcomponents-icons/dist/add.js'
import type { Task, TaskStatus } from './types'

const TAG_DESIGN: Record<TaskStatus, 'Positive' | 'Critical' | 'Information' | 'Neutral'> = {
  open: 'Neutral',
  in_progress: 'Information',
  review: 'Critical',
  completed: 'Positive',
}

const STATUS_LABEL: Record<TaskStatus, string> = {
  open: 'Open',
  in_progress: 'In Progress',
  review: 'Review',
  completed: 'Completed',
}

function TasksTable() {
  const [tasks, setTasks] = useState<Task[]>([])
  const [loading, setLoading] = useState(true)
  const [error, setError] = useState<string | null>(null)

  useEffect(() => {
    fetch('/odata/v4/tasks/Tasks')
      .then((res) => {
        if (!res.ok) throw new Error(`HTTP ${res.status}`)
        return res.json()
      })
      .then((data: { value: Task[] }) => setTasks(data.value ?? []))
      .catch((err: unknown) => setError(String(err)))
      .finally(() => setLoading(false))
  }, [])

  if (loading) {
    return <BusyIndicator active size="L" style={{ marginTop: '2rem' }} />
  }

  if (error) {
    return (
      <MessageStrip design="Negative" hideCloseButton style={{ marginBottom: '1rem' }}>
        Failed to load tasks: {error}
      </MessageStrip>
    )
  }

  if (tasks.length === 0) {
    return (
      <MessageStrip design="Information" hideCloseButton>
        No tasks found. Start by adding a new task.
      </MessageStrip>
    )
  }

  return (
    <Table
      headerRow={
        <TableHeaderRow sticky>
          <TableHeaderCell>Title</TableHeaderCell>
          <TableHeaderCell>Description</TableHeaderCell>
          <TableHeaderCell>Due Date</TableHeaderCell>
          <TableHeaderCell>Status</TableHeaderCell>
        </TableHeaderRow>
      }
    >
      {tasks.map((task) => (
        <TableRow key={task.ID}>
          <TableCell>{task.title}</TableCell>
          <TableCell>{task.description}</TableCell>
          <TableCell>{task.dueDate ?? '—'}</TableCell>
          <TableCell>
            <Tag design={TAG_DESIGN[task.status]}>
              {STATUS_LABEL[task.status] ?? task.status}
            </Tag>
          </TableCell>
        </TableRow>
      ))}
    </Table>
  )
}

export default function App() {
  return (
    <ThemeProvider>
      <ShellBar primaryTitle="SAP Tasks" secondaryTitle="Task Management">
        <ShellBarItem icon={addIcon} text="New Task" />
      </ShellBar>

      <FlexBox direction="Column" style={{ padding: '1rem 2rem' }}>
        <Title level="H3" style={{ marginBottom: '1rem' }}>
          Tasks
        </Title>
        <TasksTable />
      </FlexBox>
    </ThemeProvider>
  )
}
