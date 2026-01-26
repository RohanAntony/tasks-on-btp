const cds = require('@sap/cds')

module.exports = cds.service.impl(async function () {
  const { Tasks, TaskHistory, TaskTags, Tags } = this.entities

  // Fields to track changes for
  const TRACKED_FIELDS = ['title', 'description', 'dueDate', 'status']

  this.before('UPDATE', Tasks, async (req) => {
    const id = req.params[0]?.ID ?? req.params[0]
    if (!id) return

    const current = await SELECT.one(Tasks).where({ ID: id })
    if (!current) return

    const entries = []
    for (const field of TRACKED_FIELDS) {
      if (field in req.data) {
        const oldValue = current[field] != null ? String(current[field]) : null
        const newValue = req.data[field] != null ? String(req.data[field]) : null
        if (oldValue !== newValue) {
          entries.push({ field, oldValue, newValue, task_ID: id })
        }
      }
    }

    if (entries.length > 0) {
      await INSERT.into(TaskHistory).entries(entries)
    }
  })

  this.after('CREATE', TaskTags, async (data) => {
    const taskId = data.task_ID
    const tagId = data.tag_ID
    if (!taskId || !tagId) return

    const tag = await SELECT.one(Tags).where({ ID: tagId })
    if (!tag) return

    await INSERT.into(TaskHistory).entries([{
      field: 'tags',
      oldValue: null,
      newValue: tag.name,
      task_ID: taskId,
    }])
  })

  this.before('DELETE', TaskTags, async (req) => {
    const taskId = req.params[0]?.task_ID
    const tagId = req.params[0]?.tag_ID
    if (!taskId || !tagId) return

    const tag = await SELECT.one(Tags).where({ ID: tagId })
    if (!tag) return

    await INSERT.into(TaskHistory).entries([{
      field: 'tags',
      oldValue: tag.name,
      newValue: null,
      task_ID: taskId,
    }])
  })
})
