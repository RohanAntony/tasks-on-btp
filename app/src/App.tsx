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
import { Dialog } from '@ui5/webcomponents-react/Dialog'
import { Bar } from '@ui5/webcomponents-react/Bar'
import { Button } from '@ui5/webcomponents-react/Button'
import { Form } from '@ui5/webcomponents-react/Form'
import { FormItem } from '@ui5/webcomponents-react/FormItem'
import { Input } from '@ui5/webcomponents-react/Input'
import { Select } from '@ui5/webcomponents-react/Select'
import { Option } from '@ui5/webcomponents-react/Option'
import { DatePicker } from '@ui5/webcomponents-react/DatePicker'
import { Label } from '@ui5/webcomponents-react/Label'
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

const EMPTY_FORM = { title: '', description: '', dueDate: '', status: 'open' as TaskStatus }

interface CreateTaskDialogProps {
  open: boolean
  onClose: () => void
  onCreated: (task: Task) => void
}

function CreateTaskDialog({ open, onClose, onCreated }: CreateTaskDialogProps) {
  const [form, setForm] = useState(EMPTY_FORM)
  const [saving, setSaving] = useState(false)
  const [error, setError] = useState<string | null>(null)

  function handleClose() {
    setForm(EMPTY_FORM)
    setError(null)
    onClose()
  }

  async function handleSubmit() {
    if (!form.title.trim()) {
      setError('Title is required.')
      return
    }
    setSaving(true)
    setError(null)
    try {
      const res = await fetch('/odata/v4/tasks/Tasks', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          title: form.title.trim(),
          description: form.description.trim() || null,
          dueDate: form.dueDate || null,
          status: form.status,
        }),
      })
      if (!res.ok) throw new Error(`HTTP ${res.status}`)
      const created: Task = await res.json()
      onCreated(created)
      handleClose()
    } catch (err: unknown) {
      setError(String(err))
    } finally {
      setSaving(false)
    }
  }

  return (
    <Dialog
      open={open}
      onClose={handleClose}
      headerText="New Task"
      footer={
        <Bar
          endContent={
            <>
              <Button design="Emphasized" onClick={handleSubmit} disabled={saving}>
                {saving ? 'Creating…' : 'Create'}
              </Button>
              <Button design="Transparent" onClick={handleClose} disabled={saving}>
                Cancel
              </Button>
            </>
          }
        />
      }
    >
      <Form style={{ minWidth: '360px', padding: '0.5rem 0' }}>
        {error && (
          <MessageStrip design="Negative" hideCloseButton style={{ marginBottom: '0.75rem' }}>
            {error}
          </MessageStrip>
        )}
        <FormItem labelContent={<Label required>Title</Label>}>
          <Input
            value={form.title}
            onInput={(e) => setForm((f) => ({ ...f, title: e.target.value }))}
            placeholder="Enter task title"
            style={{ width: '100%' }}
          />
        </FormItem>
        <FormItem labelContent={<Label>Description</Label>}>
          <Input
            value={form.description}
            onInput={(e) => setForm((f) => ({ ...f, description: e.target.value }))}
            placeholder="Optional description"
            style={{ width: '100%' }}
          />
        </FormItem>
        <FormItem labelContent={<Label>Due Date</Label>}>
          <DatePicker
            value={form.dueDate}
            onChange={(e) => setForm((f) => ({ ...f, dueDate: e.detail.value ?? '' }))}
            style={{ width: '100%' }}
          />
        </FormItem>
        <FormItem labelContent={<Label>Status</Label>}>
          <Select
            onChange={(e) =>
              setForm((f) => ({ ...f, status: e.detail.selectedOption.dataset.id as TaskStatus }))
            }
            style={{ width: '100%' }}
          >
            {(Object.keys(STATUS_LABEL) as TaskStatus[]).map((s) => (
              <Option key={s} data-id={s} selected={form.status === s}>
                {STATUS_LABEL[s]}
              </Option>
            ))}
          </Select>
        </FormItem>
      </Form>
    </Dialog>
  )
}

export default function App() {
  const [tasks, setTasks] = useState<Task[]>([])
  const [loading, setLoading] = useState(true)
  const [fetchError, setFetchError] = useState<string | null>(null)
  const [dialogOpen, setDialogOpen] = useState(false)

  useEffect(() => {
    fetch('/odata/v4/tasks/Tasks')
      .then((res) => {
        if (!res.ok) throw new Error(`HTTP ${res.status}`)
        return res.json()
      })
      .then((data: { value: Task[] }) => setTasks(data.value ?? []))
      .catch((err: unknown) => setFetchError(String(err)))
      .finally(() => setLoading(false))
  }, [])

  return (
    <ThemeProvider>
      <ShellBar primaryTitle="SAP Tasks" secondaryTitle="Task Management">
        <ShellBarItem
          icon={addIcon}
          text="New Task"
          onClick={() => setDialogOpen(true)}
        />
      </ShellBar>

      <FlexBox direction="Column" style={{ padding: '1rem 2rem' }}>
        <Title level="H3" style={{ marginBottom: '1rem' }}>
          Tasks
        </Title>

        {loading && <BusyIndicator active size="L" style={{ marginTop: '2rem' }} />}

        {fetchError && (
          <MessageStrip design="Negative" hideCloseButton style={{ marginBottom: '1rem' }}>
            Failed to load tasks: {fetchError}
          </MessageStrip>
        )}

        {!loading && !fetchError && tasks.length === 0 && (
          <MessageStrip design="Information" hideCloseButton>
            No tasks found. Start by adding a new task.
          </MessageStrip>
        )}

        {!loading && !fetchError && tasks.length > 0 && (
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
        )}
      </FlexBox>

      <CreateTaskDialog
        open={dialogOpen}
        onClose={() => setDialogOpen(false)}
        onCreated={(task) => setTasks((prev) => [...prev, task])}
      />
    </ThemeProvider>
  )
}
