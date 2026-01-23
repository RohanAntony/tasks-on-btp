const cds = require('@sap/cds')

module.exports = cds.service.impl(async function () {
  const { Tasks, TaskHistory } = this.entities

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
})
