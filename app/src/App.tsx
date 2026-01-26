import '@ui5/webcomponents-react/dist/Assets.js'
import { useEffect, useRef, useState } from 'react'
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
import { TextArea } from '@ui5/webcomponents-react/TextArea'
import { Select } from '@ui5/webcomponents-react/Select'
import { Option } from '@ui5/webcomponents-react/Option'
import { DatePicker } from '@ui5/webcomponents-react/DatePicker'
import { Label } from '@ui5/webcomponents-react/Label'
import addIcon from '@ui5/webcomponents-icons/dist/add.js'
import editIcon from '@ui5/webcomponents-icons/dist/edit.js'
import historyIcon from '@ui5/webcomponents-icons/dist/history.js'
import declineIcon from '@ui5/webcomponents-icons/dist/decline.js'
import type { InputDomRef } from '@ui5/webcomponents-react'
import type { Task, TaskHistory, TaskStatus, Tag as TagType } from './types'

const DATE_FORMAT: Intl.DateTimeFormatOptions = { day: '2-digit', month: 'short', year: 'numeric' }
const formatDate = (value: string | null | undefined) =>
  value ? new Date(value).toLocaleDateString('en-GB', DATE_FORMAT) : '—'

const DATETIME_FORMAT: Intl.DateTimeFormatOptions = { ...DATE_FORMAT, hour: '2-digit', minute: '2-digit', second: '2-digit' }
const formatDateTime = (value: string | null | undefined) =>
  value ? new Date(value).toLocaleString('en-GB', DATETIME_FORMAT) : '—'

const STATUS_DESIGN: Record<TaskStatus, 'Positive' | 'Critical' | 'Information' | 'Neutral'> = {
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

function tagColorScheme(name: string): string {
  let hash = 0
  for (let i = 0; i < name.length; i++) {
    hash = (hash * 31 + name.charCodeAt(i)) >>> 0
  }
  return String((hash % 10) + 1)
}

const EMPTY_FORM = { title: '', description: '', dueDate: '', status: 'open' as TaskStatus }
type FormValues = typeof EMPTY_FORM

// ─── TaskDialog ───────────────────────────────────────────────────────────────

interface TaskDialogProps {
  open: boolean
  editTask: Task | null
  onClose: () => void
  onCreated: (task: Task) => void
  onUpdated: (task: Task) => void
}

function TaskDialog({ open, editTask, onClose, onCreated, onUpdated }: TaskDialogProps) {
  const isEdit = editTask !== null
  const [form, setForm] = useState<FormValues>(EMPTY_FORM)
  const [tags, setTags] = useState<TagType[]>([])   // tags on this task
  const [tagInput, setTagInput] = useState('')
  const [saving, setSaving] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const tagInputRef = useRef<InputDomRef>(null)

  useEffect(() => {
    if (open) {
      setForm(
        isEdit
          ? { title: editTask.title ?? '', description: editTask.description ?? '', dueDate: editTask.dueDate ?? '', status: editTask.status }
          : EMPTY_FORM,
      )
      setTags(isEdit ? (editTask.tags ?? []).map((tt) => tt.tag) : [])
      setTagInput('')
      setError(null)
    }
  }, [open])

  function addTag() {
    const name = tagInput.trim()
    if (!name || tags.some((t) => t.name.toLowerCase() === name.toLowerCase())) {
      setTagInput('')
      return
    }
    setTags((prev) => [...prev, { ID: '', name }])
    setTagInput('')
    tagInputRef.current?.focus()
  }

  function removeTag(name: string) {
    setTags((prev) => prev.filter((t) => t.name !== name))
  }

  function handleClose() {
    setError(null)
    onClose()
  }

  async function handleSubmit() {
    if (!form.title.trim()) { setError('Title is required.'); return }
    setSaving(true)
    setError(null)
    try {
      let savedTask: Task

      if (isEdit) {
        const res = await fetch(`/odata/v4/tasks/Tasks(${editTask.ID})`, {
          method: 'PATCH',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ title: form.title.trim(), description: form.description.trim() || null, dueDate: form.dueDate || null, status: form.status }),
        })
        if (!res.ok) throw new Error(`HTTP ${res.status}`)
        savedTask = { ...editTask, ...form, title: form.title.trim(), description: form.description.trim() || '' }
      } else {
        const res = await fetch('/odata/v4/tasks/Tasks', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ title: form.title.trim(), description: form.description.trim() || null, dueDate: form.dueDate || null, status: form.status }),
        })
        if (!res.ok) throw new Error(`HTTP ${res.status}`)
        savedTask = await res.json()
      }

      // Sync tags: resolve each tag name to an ID (create if needed), then replace task tags
      const resolvedTags = await Promise.all(
        tags.map(async (t) => {
          if (t.ID) return t
          // Try to find existing tag by name
          const searchRes = await fetch(`/odata/v4/tasks/Tags?$filter=name eq '${encodeURIComponent(t.name)}'`)
          const searchData: { value: TagType[] } = await searchRes.json()
          if (searchData.value.length > 0) return searchData.value[0]
          // Create new tag
          const createRes = await fetch('/odata/v4/tasks/Tags', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ name: t.name }),
          })
          return (await createRes.json()) as TagType
        }),
      )

      // Delete all existing task-tag links then re-create
      const existingLinks = editTask?.tags ?? []
      await Promise.all(
        existingLinks.map((tt) =>
          fetch(`/odata/v4/tasks/TaskTags(task_ID=${savedTask.ID},tag_ID=${tt.tag.ID})`, { method: 'DELETE' }),
        ),
      )
      await Promise.all(
        resolvedTags.map((tag) =>
          fetch('/odata/v4/tasks/TaskTags', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ task_ID: savedTask.ID, tag_ID: tag.ID }),
          }),
        ),
      )

      savedTask.tags = resolvedTags.map((tag) => ({ tag_ID: tag.ID, tag }))
      isEdit ? onUpdated(savedTask) : onCreated(savedTask)
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
      headerText={isEdit ? 'Edit Task' : 'New Task'}
      footer={
        <Bar endContent={
          <>
            <Button design="Emphasized" onClick={handleSubmit} disabled={saving}>
              {saving ? (isEdit ? 'Saving…' : 'Creating…') : isEdit ? 'Save' : 'Create'}
            </Button>
            <Button design="Transparent" onClick={handleClose} disabled={saving}>Cancel</Button>
          </>
        } />
      }
    >
      <Form style={{ minWidth: '400px', padding: '0.5rem 0' }}>
        {error && (
          <MessageStrip design="Negative" hideCloseButton style={{ marginBottom: '0.75rem' }}>{error}</MessageStrip>
        )}
        <FormItem labelContent={<Label required>Title</Label>}>
          <Input value={form.title} onInput={(e) => setForm((f) => ({ ...f, title: e.target.value }))} placeholder="Enter task title" style={{ width: '100%' }} />
        </FormItem>
        <FormItem labelContent={<Label>Description</Label>}>
          <TextArea value={form.description} onInput={(e) => setForm((f) => ({ ...f, description: e.target.value }))} placeholder="Optional description" rows={4} style={{ width: '100%' }} />
        </FormItem>
        <FormItem labelContent={<Label>Due Date</Label>}>
          <DatePicker value={form.dueDate} onChange={(e) => setForm((f) => ({ ...f, dueDate: e.detail.value ?? '' }))} style={{ width: '100%' }} />
        </FormItem>
        <FormItem labelContent={<Label>Status</Label>}>
          <Select onChange={(e) => setForm((f) => ({ ...f, status: e.detail.selectedOption.dataset.id as TaskStatus }))} style={{ width: '100%' }}>
            {(Object.keys(STATUS_LABEL) as TaskStatus[]).map((s) => (
              <Option key={s} data-id={s} selected={form.status === s}>{STATUS_LABEL[s]}</Option>
            ))}
          </Select>
        </FormItem>
        <FormItem labelContent={<Label>Tags</Label>}>
          <FlexBox direction="Column" style={{ width: '100%', gap: '0.5rem' }}>
            <FlexBox style={{ gap: '0.25rem', flexWrap: 'wrap', alignItems: 'center' }}>
              {tags.map((t) => (
                <FlexBox key={t.name} style={{ alignItems: 'center', gap: '0.125rem' }}>
                  <Tag colorScheme={tagColorScheme(t.name)}>{t.name}</Tag>
                  <Button icon={declineIcon} design="Transparent" tooltip={`Remove ${t.name}`} onClick={() => removeTag(t.name)} />
                </FlexBox>
              ))}
            </FlexBox>
            <FlexBox style={{ gap: '0.5rem' }}>
              <Input
                ref={tagInputRef}
                value={tagInput}
                onInput={(e) => setTagInput(e.target.value)}
                onKeyDown={(e) => { if (e.key === 'Enter') { e.preventDefault(); addTag() } }}
                placeholder="Add a tag…"
                style={{ flex: 1 }}
              />
              <Button icon={addIcon} design="Default" onClick={addTag}>Add</Button>
            </FlexBox>
          </FlexBox>
        </FormItem>
      </Form>
    </Dialog>
  )
}

// ─── TaskHistoryDialog ────────────────────────────────────────────────────────

interface TaskHistoryDialogProps {
  task: Task | null
  onClose: () => void
}

function TaskHistoryDialog({ task, onClose }: TaskHistoryDialogProps) {
  const [history, setHistory] = useState<TaskHistory[]>([])
  const [loading, setLoading] = useState(false)

  useEffect(() => {
    if (!task) return
    setLoading(true)
    fetch(`/odata/v4/tasks/Tasks(${task.ID})/history?$orderby=createdAt desc`)
      .then((res) => res.json())
      .then((data: { value: TaskHistory[] }) => setHistory(data.value ?? []))
      .finally(() => setLoading(false))
  }, [task])

  return (
    <Dialog
      open={!!task}
      onClose={onClose}
      headerText={task ? `History — ${task.title}` : 'History'}
      footer={<Bar endContent={<Button design="Transparent" onClick={onClose}>Close</Button>} />}
    >
      <div style={{ minWidth: '480px', padding: '0.5rem 0' }}>
        {loading && <BusyIndicator active size="M" />}
        {!loading && history.length === 0 && (
          <MessageStrip design="Information" hideCloseButton>No changes recorded yet.</MessageStrip>
        )}
        {!loading && history.length > 0 && (
          <Table
            headerRow={
              <TableHeaderRow>
                <TableHeaderCell>Field</TableHeaderCell>
                <TableHeaderCell>Old Value</TableHeaderCell>
                <TableHeaderCell>New Value</TableHeaderCell>
                <TableHeaderCell>Changed At</TableHeaderCell>
              </TableHeaderRow>
            }
          >
            {history.map((entry) => (
              <TableRow key={entry.ID}>
                <TableCell>{entry.field}</TableCell>
                <TableCell>{entry.field === 'dueDate' ? formatDate(entry.oldValue) : (entry.oldValue ?? '—')}</TableCell>
                <TableCell>{entry.field === 'dueDate' ? formatDate(entry.newValue) : (entry.newValue ?? '—')}</TableCell>
                <TableCell>{entry.createdAt ? formatDateTime(entry.createdAt) : '—'}</TableCell>
              </TableRow>
            ))}
          </Table>
        )}
      </div>
    </Dialog>
  )
}

// ─── App ──────────────────────────────────────────────────────────────────────

export default function App() {
  const [tasks, setTasks] = useState<Task[]>([])
  const [loading, setLoading] = useState(true)
  const [fetchError, setFetchError] = useState<string | null>(null)
  const [dialogOpen, setDialogOpen] = useState(false)
  const [editTask, setEditTask] = useState<Task | null>(null)
  const [historyTask, setHistoryTask] = useState<Task | null>(null)

  useEffect(() => {
    fetch('/odata/v4/tasks/Tasks?$expand=tags($expand=tag)')
      .then((res) => {
        if (!res.ok) throw new Error(`HTTP ${res.status}`)
        return res.json()
      })
      .then((data: { value: Task[] }) => setTasks(data.value ?? []))
      .catch((err: unknown) => setFetchError(String(err)))
      .finally(() => setLoading(false))
  }, [])

  function openCreate() { setEditTask(null); setDialogOpen(true) }
  function openEdit(task: Task) { setEditTask(task); setDialogOpen(true) }
  function handleClose() { setDialogOpen(false); setEditTask(null) }

  return (
    <ThemeProvider>
      <ShellBar primaryTitle="SAP Tasks" secondaryTitle="Task Management">
        <ShellBarItem icon={addIcon} text="New Task" onClick={openCreate} />
      </ShellBar>

      <FlexBox direction="Column" style={{ padding: '1rem 2rem' }}>
        <Title level="H3" style={{ marginBottom: '1rem' }}>Tasks</Title>

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
                <TableHeaderCell>Tags</TableHeaderCell>
                <TableHeaderCell>Created</TableHeaderCell>
                <TableHeaderCell>Due Date</TableHeaderCell>
                <TableHeaderCell>Status</TableHeaderCell>
                <TableHeaderCell />
              </TableHeaderRow>
            }
          >
            {tasks.map((task) => (
              <TableRow key={task.ID}>
                <TableCell>{task.title}</TableCell>
                <TableCell>{task.description}</TableCell>
                <TableCell>
                  <FlexBox style={{ gap: '0.25rem', flexWrap: 'wrap' }}>
                    {(task.tags ?? []).map((tt) => (
                      <Tag key={tt.tag.ID} colorScheme={tagColorScheme(tt.tag.name)} design="Set2">{tt.tag.name}</Tag>
                    ))}
                  </FlexBox>
                </TableCell>
                <TableCell>{formatDate(task.createdAt)}</TableCell>
                <TableCell>{formatDate(task.dueDate)}</TableCell>
                <TableCell>
                  <Tag design={STATUS_DESIGN[task.status]}>
                    {STATUS_LABEL[task.status] ?? task.status}
                  </Tag>
                </TableCell>
                <TableCell>
                  <Button icon={editIcon} design="Transparent" tooltip="Edit task" onClick={() => openEdit(task)} />
                  <Button icon={historyIcon} design="Transparent" tooltip="View history" onClick={() => setHistoryTask(task)} />
                </TableCell>
              </TableRow>
            ))}
          </Table>
        )}
      </FlexBox>

      <TaskDialog
        open={dialogOpen}
        editTask={editTask}
        onClose={handleClose}
        onCreated={(task) => setTasks((prev) => [...prev, task])}
        onUpdated={(updated) => setTasks((prev) => prev.map((t) => (t.ID === updated.ID ? updated : t)))}
      />

      <TaskHistoryDialog task={historyTask} onClose={() => setHistoryTask(null)} />
    </ThemeProvider>
  )
}
