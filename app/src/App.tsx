import '@ui5/webcomponents-react/dist/Assets.js'
import { setTheme } from '@ui5/webcomponents-base/dist/config/Theme.js'
import React, { useEffect, useRef, useState } from 'react'
import { ThemeProvider } from '@ui5/webcomponents-react/ThemeProvider'
import { ShellBar } from '@ui5/webcomponents-react/ShellBar'
import { ShellBarItem } from '@ui5/webcomponents-react/ShellBarItem'
import { FlexibleColumnLayout } from '@ui5/webcomponents-react/FlexibleColumnLayout'
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
import { Text } from '@ui5/webcomponents-react/Text'
import FCLLayout from '@ui5/webcomponents-fiori/dist/types/FCLLayout.js'
import { ObjectPage } from '@ui5/webcomponents-react/ObjectPage'
import { ObjectPageSection } from '@ui5/webcomponents-react/ObjectPageSection'
import { ObjectPageTitle } from '@ui5/webcomponents-react/ObjectPageTitle'

import addIcon from '@ui5/webcomponents-icons/dist/add.js'
import editIcon from '@ui5/webcomponents-icons/dist/edit.js'
import declineIcon from '@ui5/webcomponents-icons/dist/decline.js'
import closeIcon from '@ui5/webcomponents-icons/dist/decline.js'
import darkModeIcon from '@ui5/webcomponents-icons/dist/dark-mode.js'
import lightModeIcon from '@ui5/webcomponents-icons/dist/light-mode.js'
import gridIcon from '@ui5/webcomponents-icons/dist/grid.js'
import tableViewIcon from '@ui5/webcomponents-icons/dist/table-view.js'
import type { InputDomRef } from '@ui5/webcomponents-react'
import type { Task, TaskHistory, TaskComment, TaskStatus, Tag as TagType } from './types'

const DATE_FORMAT: Intl.DateTimeFormatOptions = { day: '2-digit', month: 'short', year: 'numeric' }
const formatDate = (value: string | null | undefined) =>
  value ? new Date(value).toLocaleDateString('en-GB', DATE_FORMAT) : '—'

const DATETIME_FORMAT: Intl.DateTimeFormatOptions = { ...DATE_FORMAT, hour: '2-digit', minute: '2-digit', second: '2-digit' }
const formatDateTime = (value: string | null | undefined) =>
  value ? new Date(value).toLocaleString('en-GB', DATETIME_FORMAT) : '—'

// ─── CSRF-aware fetch ─────────────────────────────────────────────────────────

let _csrfToken: string | null = null

async function getCsrfToken(): Promise<string | null> {
  if (_csrfToken) return _csrfToken
  const res = await fetch('/odata/v4/tasks/', { headers: { 'x-csrf-token': 'fetch' } })
  _csrfToken = res.headers.get('x-csrf-token')
  return _csrfToken
}

async function mutate(url: string, method: string, body?: object): Promise<Response> {
  const token = await getCsrfToken()
  const headers: Record<string, string> = { 'Content-Type': 'application/json' }
  if (token) headers['x-csrf-token'] = token
  return fetch(url, { method, headers, body: body ? JSON.stringify(body) : undefined })
}

const STATUS_DESIGN: Record<TaskStatus, 'Positive' | 'Critical' | 'Information' | 'Neutral' | 'Negative'> = {
  open: 'Negative',
  in_progress: 'Critical',
  review: 'Information',
  completed: 'Positive',
}

const STATUS_LABEL: Record<TaskStatus, string> = {
  open: 'Open',
  in_progress: 'In Progress',
  review: 'Review',
  completed: 'Completed',
}

const HISTORY_FIELD_LABEL: Record<string, string> = {
  title: 'Title',
  description: 'Description',
  dueDate: 'Due Date',
  status: 'Status',
  tags: 'Tags',
}

function formatHistoryValue(field: string, value: string | null): string {
  if (value == null) return '—'
  if (field === 'dueDate') return formatDate(value)
  if (field === 'status') return STATUS_LABEL[value as TaskStatus] ?? value
  return value
}

function tagTextColor(hex: string): string {
  const r = parseInt(hex.slice(1, 3), 16)
  const g = parseInt(hex.slice(3, 5), 16)
  const b = parseInt(hex.slice(5, 7), 16)
  const luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255
  return luminance > 0.55 ? '#3a3a3a' : '#ffffff'
}

function ColorTag({ name, color }: { name: string; color: string | null }) {
  const bg = color ?? '#e0e0e0'
  const fg = color ? tagTextColor(color) : '#3a3a3a'
  return (
    <span style={{
      display: 'inline-flex',
      alignItems: 'center',
      padding: '0.125rem 0.5rem',
      borderRadius: '0.75rem',
      fontSize: 'var(--sapFontSmallSize)',
      fontFamily: 'var(--sapFontFamily)',
      backgroundColor: bg,
      color: fg,
      border: `1px solid ${bg}`,
      whiteSpace: 'nowrap',
    }}>
      {name}
    </span>
  )
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
  const [tags, setTags] = useState<TagType[]>([])
  const [allTags, setAllTags] = useState<TagType[]>([])
  const [tagInput, setTagInput] = useState('')
  const [tagColor, setTagColor] = useState('#b3d9ff')
  const [tagTab, setTagTab] = useState<'existing' | 'new'>('existing')
  const [selectedExistingTagId, setSelectedExistingTagId] = useState('')
  const [saving, setSaving] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const tagInputRef = useRef<InputDomRef>(null)

  useEffect(() => {
    if (!open) return
    fetch('/odata/v4/tasks/Tags?$orderby=name')
      .then((res) => res.json())
      .then((data: { value: TagType[] }) => setAllTags((data.value ?? []).filter((t) => t != null && t.name != null)))
  }, [open])

  useEffect(() => {
    if (open) {
      setForm(
        isEdit
          ? { title: editTask.title ?? '', description: editTask.description ?? '', dueDate: editTask.dueDate ?? '', status: editTask.status }
          : EMPTY_FORM,
      )
      setTags(isEdit ? (editTask.tags ?? []).filter((tt) => tt.tag != null).map((tt) => tt.tag) : [])
      setTagInput('')
      setTagColor('#b3d9ff')
      setTagTab('existing')
      setSelectedExistingTagId('')
      setError(null)
    }
  }, [open])

  function selectExistingTag(tag: TagType) {
    if (tags.some((t) => t.ID === tag.ID)) return
    setTags((prev) => [...prev, tag])
  }

  function addNewTag() {
    const name = tagInput.trim()
    if (!name || tags.some((t) => t.name.toLowerCase() === name.toLowerCase())) { setTagInput(''); return }
    setTags((prev) => [...prev, { ID: '', name, color: tagColor }])
    setTagInput('')
    setTagColor('#b3d9ff')
    tagInputRef.current?.focus()
  }

  function removeTag(name: string) { setTags((prev) => prev.filter((t) => t.name !== name)) }
  function handleClose() { setError(null); onClose() }

  async function handleSubmit() {
    if (!form.title.trim()) { setError('Title is required.'); return }
    setSaving(true); setError(null)
    try {
      let savedTask: Task
      if (isEdit) {
        const res = await mutate(`/odata/v4/tasks/Tasks('${editTask.ID}')`, 'PATCH',
          { title: form.title.trim(), description: form.description.trim() || null, dueDate: form.dueDate || null, status: form.status })
        if (!res.ok) throw new Error(`HTTP ${res.status}`)
        savedTask = { ...editTask, ...form, title: form.title.trim(), description: form.description.trim() || '' }
      } else {
        const res = await mutate('/odata/v4/tasks/Tasks', 'POST',
          { title: form.title.trim(), description: form.description.trim() || null, dueDate: form.dueDate || null, status: form.status })
        if (!res.ok) throw new Error(`HTTP ${res.status}`)
        savedTask = await res.json()
      }

      const resolvedTags = await Promise.all(
        tags.map(async (t) => {
          if (t.ID) return t
          const searchRes = await fetch(`/odata/v4/tasks/Tags?$filter=name eq '${encodeURIComponent(t.name)}'`)
          const searchData: { value: TagType[] } = await searchRes.json()
          if (searchData.value.length > 0) return searchData.value[0]
          const createRes = await mutate('/odata/v4/tasks/Tags', 'POST', { name: t.name, color: t.color ?? null })
          const created = (await createRes.json()) as TagType
          setAllTags((prev) => [...prev, created].filter((t) => t != null && t.name != null))
          return created
        }),
      )

      const existingLinks = (editTask?.tags ?? []).filter((tt) => tt.tag != null)
      await Promise.all(existingLinks.map((tt) => mutate(`/odata/v4/tasks/TaskTags(task_ID='${savedTask.ID}',tag_ID='${tt.tag.ID}')`, 'DELETE')))
      await Promise.all(resolvedTags.map((tag) => mutate('/odata/v4/tasks/TaskTags', 'POST', { task_ID: savedTask.ID, tag_ID: tag.ID })))

      savedTask.tags = resolvedTags.map((tag) => ({ tag_ID: tag.ID, tag }))
      isEdit ? onUpdated(savedTask) : onCreated(savedTask)
      handleClose()
    } catch (err: unknown) {
      setError(String(err))
    } finally {
      setSaving(false)
    }
  }

  const availableTags = allTags.filter((t) => t != null && t.name != null && !tags.some((sel) => sel.ID === t.ID))

  useEffect(() => {
    if (tagTab === 'existing' && availableTags.length > 0 && !availableTags.find((t) => t.ID === selectedExistingTagId)) {
      setSelectedExistingTagId(availableTags[0].ID)
    }
  }, [availableTags, tagTab])

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
      <Form style={{ minWidth: '460px', padding: '0.5rem 0' }}>
        {error && <MessageStrip design="Negative" hideCloseButton style={{ marginBottom: '0.75rem' }}>{error}</MessageStrip>}
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

            {/* Selected tags */}
            {tags.length > 0 && (
              <FlexBox style={{ gap: '0.25rem', flexWrap: 'wrap', alignItems: 'center' }}>
                {tags.map((t) => (
                  <FlexBox key={t.name} style={{ alignItems: 'center', gap: '0.125rem' }}>
                    <ColorTag name={t.name} color={t.color ?? null} />
                    <Button icon={declineIcon} design="Transparent" tooltip={`Remove ${t.name}`} onClick={() => removeTag(t.name)} />
                  </FlexBox>
                ))}
              </FlexBox>
            )}

            {/* Tab bar */}
            <FlexBox style={{ borderBottom: '1px solid var(--sapGroup_TitleBorderColor)' }}>
              {(['existing', 'new'] as const).map((tab) => (
                <button
                  key={tab}
                  onClick={() => setTagTab(tab)}
                  style={{
                    background: 'none',
                    border: 'none',
                    borderBottom: tagTab === tab ? '2px solid var(--sapSelectedColor)' : '2px solid transparent',
                    color: tagTab === tab ? 'var(--sapSelectedColor)' : 'var(--sapContent_LabelColor)',
                    fontFamily: 'var(--sapFontFamily)',
                    fontSize: 'var(--sapFontSize)',
                    fontWeight: tagTab === tab ? 'bold' : 'normal',
                    padding: '0.35rem 0.75rem',
                    cursor: 'pointer',
                    marginBottom: '-1px',
                  }}
                >
                  {tab === 'existing' ? 'Select Existing' : 'Create New'}
                </button>
              ))}
            </FlexBox>

            {/* Pick existing tag */}
            {tagTab === 'existing' && (
              <FlexBox style={{ gap: '0.5rem', alignItems: 'center' }}>
                <Select
                  onChange={(e) => setSelectedExistingTagId(e.detail.selectedOption.dataset.id ?? '')}
                  style={{ flex: 1 }}
                >
                  {availableTags.map((t) => (
                    <Option key={t.ID} data-id={t.ID}>{t.name}</Option>
                  ))}
                </Select>
                <Button
                  icon={addIcon}
                  design="Default"
                  disabled={!selectedExistingTagId}
                  onClick={() => {
                    const tag = allTags.find((t) => t.ID === selectedExistingTagId)
                    if (tag) { selectExistingTag(tag); setSelectedExistingTagId('') }
                  }}
                >
                  Add
                </Button>
              </FlexBox>
            )}

            {/* Create new tag */}
            {tagTab === 'new' && (
              <FlexBox style={{ gap: '0.5rem', alignItems: 'center' }}>
                <Input
                  ref={tagInputRef}
                  value={tagInput}
                  onInput={(e) => setTagInput(e.target.value)}
                  onKeyDown={(e) => { if (e.key === 'Enter') { e.preventDefault(); addNewTag() } }}
                  placeholder="New tag name…"
                  style={{ flex: 1 }}
                />
                <input
                  type="color"
                  value={tagColor}
                  onChange={(e) => setTagColor(e.target.value)}
                  title="Tag color"
                  style={{ width: '2.25rem', height: '2.25rem', border: '1px solid var(--sapField_BorderColor)', padding: '0.2rem', borderRadius: '0.25rem', cursor: 'pointer', background: 'none', boxSizing: 'border-box' }}
                />
                <Button icon={addIcon} design="Default" onClick={addNewTag}>Add</Button>
              </FlexBox>
            )}

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
    fetch(`/odata/v4/tasks/Tasks('${task.ID}')/history?$orderby=createdAt desc`)
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

// ─── TaskDetailPanel ──────────────────────────────────────────────────────────

interface TaskDetailPanelProps {
  task: Task
  refreshKey: number
  onEdit: () => void
  onClose: () => void
}

function TaskDetailPanel({ task, refreshKey, onEdit, onClose }: TaskDetailPanelProps) {
  const [history, setHistory] = useState<TaskHistory[]>([])
  const [historyLoading, setHistoryLoading] = useState(false)
  const [comments, setComments] = useState<TaskComment[]>(task.comments ?? [])
  const [newComment, setNewComment] = useState('')
  const [commentSaving, setCommentSaving] = useState(false)

  useEffect(() => {
    setHistoryLoading(true)
    fetch(`/odata/v4/tasks/Tasks('${task.ID}')/history?$orderby=createdAt desc`)
      .then((res) => res.json())
      .then((data: { value: TaskHistory[] }) => setHistory(data.value ?? []))
      .finally(() => setHistoryLoading(false))
  }, [task.ID, refreshKey])

  useEffect(() => {
    setComments(task.comments ?? [])
  }, [task.ID])

  async function addComment() {
    const content = newComment.trim()
    if (!content) return
    setCommentSaving(true)
    try {
      const res = await mutate('/odata/v4/tasks/TaskComments', 'POST', { task_ID: task.ID, content })
      if (!res.ok) throw new Error(`HTTP ${res.status}`)
      const created: TaskComment = await res.json()
      setComments((prev) => [...prev, created])
      setNewComment('')
    } finally {
      setCommentSaving(false)
    }
  }

  const gridStyle: React.CSSProperties = {
    display: 'grid',
    gridTemplateColumns: 'auto 1fr',
    gap: '0.6rem 1.25rem',
    alignItems: 'center',
  }

  return (
    <ObjectPage
      style={{ height: '100%', borderLeft: '1px solid var(--sapGroup_TitleBorderColor)' }}
      titleArea={
        <ObjectPageTitle
          header={task.title}
          actionsBar={
            <FlexBox style={{ gap: '0.25rem' }}>
              <Button icon={editIcon} design="Transparent" tooltip="Edit task" onClick={onEdit} />
              <Button icon={closeIcon} design="Transparent" tooltip="Close" onClick={onClose} />
            </FlexBox>
          }
        />
      }
    >
      <ObjectPageSection id="details" titleText="Details">
          <div style={gridStyle}>
            <Label>Status</Label>
            <Tag design={STATUS_DESIGN[task.status]}>{STATUS_LABEL[task.status]}</Tag>

            <Label>Created</Label>
            <span style={{ fontSize: 'var(--sapFontSize)', color: 'var(--sapTextColor)' }}>{formatDate(task.createdAt)}</span>

            <Label>Due Date</Label>
            <span style={{ fontSize: 'var(--sapFontSize)', color: 'var(--sapTextColor)' }}>{formatDate(task.dueDate)}</span>

            {(task.tags ?? []).filter((tt) => tt.tag != null).length > 0 && (
              <>
                <Label>Tags</Label>
                <FlexBox style={{ gap: '0.25rem', flexWrap: 'wrap' }}>
                  {(task.tags ?? []).filter((tt) => tt.tag != null).map((tt) => (
                    <ColorTag key={tt.tag.ID} name={tt.tag.name} color={tt.tag.color ?? null} />
                  ))}
                </FlexBox>
              </>
            )}
          </div>
      </ObjectPageSection>

      {!!task.description && (
        <ObjectPageSection id="description" titleText="Description">
          <Text style={{ whiteSpace: 'pre-wrap' }}>{task.description}</Text>
        </ObjectPageSection>
      )}

      <ObjectPageSection id="comments" titleText="Comments">
          <FlexBox direction="Column" style={{ gap: '0.75rem' }}>
            {comments.length === 0 && (
              <Text style={{ color: 'var(--sapNeutralColor)' }}>No comments yet.</Text>
            )}
            {comments.map((c) => (
              <div key={c.ID} style={{ borderLeft: '3px solid var(--sapList_HeaderBackground)', paddingLeft: '0.75rem' }}>
                <div style={{ fontSize: 'var(--sapFontSmallSize)', color: 'var(--sapContent_LabelColor)', marginBottom: '0.25rem' }}>
                  {formatDateTime(c.createdAt)}
                </div>
                <Text style={{ whiteSpace: 'pre-wrap' }}>{c.content}</Text>
              </div>
            ))}
            <FlexBox style={{ gap: '0.5rem', alignItems: 'flex-end' }}>
              <TextArea
                value={newComment}
                onInput={(e) => setNewComment(e.target.value)}
                placeholder="Add a comment…"
                rows={2}
                style={{ flex: 1 }}
              />
              <Button design="Emphasized" onClick={addComment} disabled={commentSaving || !newComment.trim()}>
                {commentSaving ? 'Adding…' : 'Add'}
              </Button>
            </FlexBox>
          </FlexBox>
      </ObjectPageSection>

      <ObjectPageSection id="history" titleText="History">
          {historyLoading && <BusyIndicator active size="S" />}
          {!historyLoading && history.length === 0 && (
            <Text style={{ color: 'var(--sapNeutralColor)' }}>No changes recorded yet.</Text>
          )}
          {!historyLoading && history.length > 0 && (
            <div style={{ display: 'grid', gridTemplateColumns: 'auto 1fr 1fr auto', gap: '0.4rem 1rem', alignItems: 'center', fontSize: 'var(--sapFontSize)', color: 'var(--sapTextColor)' }}>
              <span style={{ fontWeight: 'bold' }}>Field</span>
              <span style={{ fontWeight: 'bold' }}>Old</span>
              <span style={{ fontWeight: 'bold' }}>New</span>
              <span style={{ fontWeight: 'bold' }}>Changed At</span>
              {history.map((entry) => (
                <React.Fragment key={entry.ID}>
                  <span>{HISTORY_FIELD_LABEL[entry.field] ?? entry.field}</span>
                  <span>{formatHistoryValue(entry.field, entry.oldValue)}</span>
                  <span>{formatHistoryValue(entry.field, entry.newValue)}</span>
                  <span>{formatDateTime(entry.createdAt)}</span>
                </React.Fragment>
              ))}
            </div>
          )}
      </ObjectPageSection>
    </ObjectPage>
  )
}

// ─── KanbanBoard ──────────────────────────────────────────────────────────────

const STATUSES: TaskStatus[] = ['open', 'in_progress', 'review', 'completed']

interface KanbanBoardProps {
  tasks: Task[]
  selectedTask: Task | null
  onSelect: (task: Task) => void
  onStatusChange: (taskId: string, newStatus: TaskStatus) => void
}

function KanbanBoard({ tasks, selectedTask, onSelect, onStatusChange }: KanbanBoardProps) {
  const [draggedTaskId, setDraggedTaskId] = useState<string | null>(null)
  const [dragOverStatus, setDragOverStatus] = useState<TaskStatus | null>(null)

  function handleDragStart(e: React.DragEvent<HTMLDivElement>, taskId: string) {
    e.dataTransfer.effectAllowed = 'move'
    e.dataTransfer.setData('text/plain', taskId)
    setDraggedTaskId(taskId)
  }

  function handleDragEnd() {
    setDraggedTaskId(null)
    setDragOverStatus(null)
  }

  function handleDragOver(e: React.DragEvent<HTMLDivElement>, status: TaskStatus) {
    e.preventDefault()
    e.dataTransfer.dropEffect = 'move'
    setDragOverStatus(status)
  }

  function handleDragLeave(e: React.DragEvent<HTMLDivElement>) {
    if (!e.currentTarget.contains(e.relatedTarget as Node)) {
      setDragOverStatus(null)
    }
  }

  function handleDrop(e: React.DragEvent<HTMLDivElement>, targetStatus: TaskStatus) {
    e.preventDefault()
    const taskId = e.dataTransfer.getData('text/plain')
    setDraggedTaskId(null)
    setDragOverStatus(null)
    const task = tasks.find((t) => t.ID === taskId)
    if (!task || task.status === targetStatus) return
    onStatusChange(taskId, targetStatus)
  }

  return (
    <div style={{
      display: 'grid',
      gridTemplateColumns: 'repeat(4, 1fr)',
      gap: '0.75rem',
      padding: '0.75rem',
      height: '100%',
      boxSizing: 'border-box',
      overflowY: 'auto',
      alignItems: 'start',
    }}>
      {STATUSES.map((status) => {
        const columnTasks = tasks.filter((t) => t.status === status)
        const isDropTarget = dragOverStatus === status
        return (
          <div
            key={status}
            onDragOver={(e) => handleDragOver(e, status)}
            onDragLeave={handleDragLeave}
            onDrop={(e) => handleDrop(e, status)}
            style={{
              display: 'flex',
              flexDirection: 'column',
              gap: '0.5rem',
              border: '1px solid var(--sapGroup_TitleBorderColor)',
              borderRadius: '0.5rem',
              overflow: 'hidden',
              background: isDropTarget ? 'var(--sapList_SelectionBackgroundColor)' : 'var(--sapList_Background)',
              outline: isDropTarget ? '2px solid var(--sapSelectedColor)' : 'none',
              outlineOffset: '-1px',
              transition: 'background 0.15s ease',
            }}
          >
            {/* Column header */}
            <div style={{
              padding: '0.5rem 0.75rem',
              borderBottom: '1px solid var(--sapGroup_TitleBorderColor)',
              display: 'flex',
              alignItems: 'center',
              gap: '0.5rem',
              background: 'var(--sapList_HeaderBackground)',
            }}>
              <Tag design={STATUS_DESIGN[status]}>{STATUS_LABEL[status]}</Tag>
              <span style={{
                marginLeft: 'auto',
                minWidth: '1.25rem',
                height: '1.25rem',
                padding: '0 0.375rem',
                borderRadius: '0.625rem',
                background: 'var(--sapNeutralBackground)',
                border: '1px solid var(--sapNeutralBorderColor)',
                fontSize: 'var(--sapFontSmallSize)',
                color: 'var(--sapNeutralTextColor)',
                display: 'inline-flex',
                alignItems: 'center',
                justifyContent: 'center',
                fontWeight: 'bold',
                boxSizing: 'border-box',
              }}>{columnTasks.length}</span>
            </div>

            {/* Cards */}
            <div style={{ display: 'flex', flexDirection: 'column', gap: '0.5rem', padding: '0.5rem' }}>
              {columnTasks.length === 0 && (
                <span style={{ fontSize: 'var(--sapFontSmallSize)', color: 'var(--sapContent_LabelColor)', textAlign: 'center', padding: '0.5rem 0' }}>
                  No tasks
                </span>
              )}
              {columnTasks.map((task) => (
                <div
                  key={task.ID}
                  draggable
                  onDragStart={(e) => handleDragStart(e, task.ID)}
                  onDragEnd={handleDragEnd}
                  onClick={() => onSelect(task)}
                  style={{
                    padding: '0.6rem 0.75rem',
                    borderRadius: '0.375rem',
                    border: `1px solid ${selectedTask?.ID === task.ID ? 'var(--sapSelectedColor)' : 'var(--sapGroup_TitleBorderColor)'}`,
                    background: selectedTask?.ID === task.ID ? 'var(--sapList_SelectionBackgroundColor)' : 'var(--sapList_Background)',
                    color: 'var(--sapTextColor)',
                    fontFamily: 'var(--sapFontFamily)',
                    fontSize: 'var(--sapFontSize)',
                    cursor: draggedTaskId === task.ID ? 'grabbing' : 'grab',
                    display: 'flex',
                    flexDirection: 'column',
                    gap: '0.35rem',
                    opacity: draggedTaskId === task.ID ? 0.4 : 1,
                    userSelect: 'none',
                  }}
                >
                  <span style={{ wordBreak: 'break-word' }}>
                    {task.title}
                  </span>
                  {task.dueDate && (
                    <span style={{ fontSize: 'var(--sapFontSmallSize)', color: 'var(--sapContent_LabelColor)' }}>
                      {formatDate(task.dueDate)}
                    </span>
                  )}
                  {(task.tags ?? []).filter((tt) => tt.tag != null).length > 0 && (
                    <div style={{ display: 'flex', flexWrap: 'wrap', gap: '0.25rem', marginTop: '0.1rem' }}>
                      {(task.tags ?? []).filter((tt) => tt.tag != null).map((tt) => (
                        <ColorTag key={tt.tag.ID} name={tt.tag.name} color={tt.tag.color ?? null} />
                      ))}
                    </div>
                  )}
                </div>
              ))}
            </div>
          </div>
        )
      })}
    </div>
  )
}

// ─── App ──────────────────────────────────────────────────────────────────────

export default function App() {
  const [tasks, setTasks] = useState<Task[]>([])
  const [loading, setLoading] = useState(true)
  const [fetchError, setFetchError] = useState<string | null>(null)
  const [selectedTask, setSelectedTask] = useState<Task | null>(null)
  const [dialogOpen, setDialogOpen] = useState(false)
  const [editTask, setEditTask] = useState<Task | null>(null)
  const [darkMode, setDarkMode] = useState(false)
  const [detailRefreshKey, setDetailRefreshKey] = useState(0)
  const [viewMode, setViewMode] = useState<'table' | 'board'>('table')

  function toggleDarkMode() {
    const next = !darkMode
    setDarkMode(next)
    setTheme(next ? 'sap_horizon_dark' : 'sap_horizon')
  }

  useEffect(() => {
    fetch('/odata/v4/tasks/Tasks?$expand=tags($expand=tag),comments($orderby=createdAt asc)')
      .then((res) => { if (!res.ok) throw new Error(`HTTP ${res.status}`); return res.json() })
      .then((data: { value: Task[] }) => setTasks(data.value ?? []))
      .catch((err: unknown) => setFetchError(String(err)))
      .finally(() => setLoading(false))
  }, [])

  function openCreate() { setEditTask(null); setDialogOpen(true) }
  function openEdit(task: Task) { setEditTask(task); setDialogOpen(true) }
  function handleClose() { setDialogOpen(false); setEditTask(null) }

  function handleUpdated(updated: Task) {
    setTasks((prev) => prev.map((t) => (t.ID === updated.ID ? updated : t)))
    if (selectedTask?.ID === updated.ID) {
      setSelectedTask(updated)
      setDetailRefreshKey((k) => k + 1)
    }
  }

  async function handleStatusChange(taskId: string, newStatus: TaskStatus) {
    const previousTasks = tasks
    setTasks((prev) => prev.map((t) => (t.ID === taskId ? { ...t, status: newStatus } : t)))
    if (selectedTask?.ID === taskId)
      setSelectedTask((prev) => prev ? { ...prev, status: newStatus } : null)
    try {
      const res = await mutate(`/odata/v4/tasks/Tasks('${taskId}')`, 'PATCH', { status: newStatus })
      if (!res.ok) throw new Error(`HTTP ${res.status}`)
    } catch {
      setTasks(previousTasks)
      if (selectedTask?.ID === taskId) {
        const rolled = previousTasks.find((t) => t.ID === taskId)
        if (rolled) setSelectedTask(rolled)
      }
    }
  }

  const startColumn = (
    <FlexBox direction="Column" style={{ height: '100%' }}>
      <FlexBox
        alignItems="Center"
        justifyContent="SpaceBetween"
        style={{ padding: '0.75rem 1rem', borderBottom: '1px solid var(--sapGroup_TitleBorderColor)' }}
      >
        <Title level="H3">Tasks</Title>
        <Button icon={addIcon} design="Emphasized" onClick={openCreate}>New Task</Button>
      </FlexBox>

      {loading && <BusyIndicator active size="L" style={{ margin: '2rem auto' }} />}

      {fetchError && (
        <MessageStrip design="Negative" hideCloseButton style={{ margin: '1rem' }}>
          Failed to load tasks: {fetchError}
        </MessageStrip>
      )}

      {!loading && !fetchError && tasks.length === 0 && (
        <MessageStrip design="Information" hideCloseButton style={{ margin: '1rem' }}>
          No tasks found. Start by adding a new task.
        </MessageStrip>
      )}

      {!loading && !fetchError && tasks.length > 0 && viewMode === 'table' && (
        <Table
          headerRow={
            <TableHeaderRow sticky>
              <TableHeaderCell>Title</TableHeaderCell>
              <TableHeaderCell>Tags</TableHeaderCell>
              <TableHeaderCell>Due Date</TableHeaderCell>
              <TableHeaderCell>Status</TableHeaderCell>
            </TableHeaderRow>
          }
        >
          {tasks.map((task) => (
            <TableRow
              key={task.ID}
              onClick={() => setSelectedTask(task)}
              style={{ cursor: 'pointer', background: selectedTask?.ID === task.ID ? 'var(--sapList_SelectionBackgroundColor)' : undefined }}
            >
              <TableCell>{task.title}</TableCell>
              <TableCell>
                <FlexBox style={{ gap: '0.25rem', flexWrap: 'wrap' }}>
                  {(task.tags ?? []).filter((tt) => tt.tag != null).map((tt) => (
                    <ColorTag key={tt.tag.ID} name={tt.tag.name} color={tt.tag.color ?? null} />
                  ))}
                </FlexBox>
              </TableCell>
              <TableCell>{formatDate(task.dueDate)}</TableCell>
              <TableCell>
                <Tag design={STATUS_DESIGN[task.status]}>{STATUS_LABEL[task.status] ?? task.status}</Tag>
              </TableCell>
            </TableRow>
          ))}
        </Table>
      )}

      {!loading && !fetchError && tasks.length > 0 && viewMode === 'board' && (
        <KanbanBoard tasks={tasks} selectedTask={selectedTask} onSelect={setSelectedTask} onStatusChange={handleStatusChange} />
      )}
    </FlexBox>
  )

  const midColumn = selectedTask ? (
    <div slot="midColumn" style={{ height: '100%' }}>
      <TaskDetailPanel
        task={selectedTask}
        refreshKey={detailRefreshKey}
        onEdit={() => openEdit(selectedTask)}
        onClose={() => setSelectedTask(null)}
      />
    </div>
  ) : undefined

  return (
    <ThemeProvider>
      <ShellBar
        primaryTitle="Task Management"
        logo={<img src="/logo.png" alt="logo" style={{ height: '1.75rem' }} />}
      >
        <ShellBarItem icon={darkMode ? lightModeIcon : darkModeIcon} text={darkMode ? 'Light Mode' : 'Dark Mode'} onClick={toggleDarkMode} />
        <ShellBarItem icon={viewMode === 'table' ? gridIcon : tableViewIcon} text={viewMode === 'table' ? 'Board View' : 'Table View'} onClick={() => setViewMode((v) => v === 'table' ? 'board' : 'table')} />
      </ShellBar>

      <FlexibleColumnLayout
        layout={selectedTask ? FCLLayout.TwoColumnsStartExpanded : FCLLayout.OneColumn}
        style={{ height: 'calc(100vh - 3rem)' }}
      >
        <div slot="startColumn" style={{ height: '100%' }}>
          {startColumn}
        </div>
        {midColumn}
      </FlexibleColumnLayout>

      <TaskDialog
        open={dialogOpen}
        editTask={editTask}
        onClose={handleClose}
        onCreated={(task) => setTasks((prev) => [...prev, task])}
        onUpdated={handleUpdated}
      />
    </ThemeProvider>
  )
}
