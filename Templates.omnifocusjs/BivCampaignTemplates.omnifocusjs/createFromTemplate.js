/* global PlugIn Form Preferences */
(() => {
  const action = new PlugIn.Action(async function (selection, sender) {
    // Access Kaitlin’s template library and preferences
    const templateLibrary = this.templateLibrary
    const preferences = new Preferences('com.KaitlinSalzke.Templates')

    // Get the folder containing template projects
    const templateFolder = await templateLibrary.getTemplateFolder()

    // If user has exactly one project selected in the template folder, use that
    let template = (
      selection &&
      selection.projects.length === 1 &&
      templateFolder.flattenedProjects.includes(selection.projects[0])
    )
      ? selection.projects[0]
      : null

    // Build the combined form (choose template if needed, gather campaign info, etc.)
    const templateForm = await generateTemplateForm(template)

    // If template was not pre-selected, show form to pick it
    if (template === null) {
      await templateForm.show('Choose Template', 'Create')
      template = templateForm.values.template
    } else {
      // If template was pre-selected, just show form for campaign fields
      await templateForm.show('Enter Campaign Info', 'Create')
    }

    // Determine destination folder/project
    const destination = await templateLibrary.getDestination(template)

    // Duplicate the template project using Kaitlin’s built-in logic
    const newProject = await templateLibrary.createFromTemplate(template, destination)

    // Replace placeholders in tasks (e.g. «Campaign Name», «Artist Name», «Release Date»)
    replacePlaceholders(newProject, {
      '«Campaign Name»': templateForm.values.campaignName,
      '«Artist Name»': templateForm.values.artistName,
      '«Release Date»': templateForm.values.releaseDate.toLocaleDateString()
    })

    // Remove tasks tagged with verticals the user did NOT select
    removeUnselectedVerticalTasks(
      newProject,
      templateForm.values.selectedVerticals // e.g. ["DSP", "Radio", "Press"]
    )

    // Parse $DEFER= / $DUE= lines in task notes, relative to the chosen release date
    applyDateOffsets(newProject, templateForm.values.releaseDate)

    // If user checked “Go to created project,” focus it in OmniFocus
    if (templateForm.values.goTo) {
      URL.fromString('omnifocus:///task/' + newProject.id.primaryKey).call(() => {})
    }
  })

  // Always allow this action to run (Kaitlin’s approach)
  action.validate = function (selection, sender) {
    return true
  }

  // ----------------------
  //     HELPER FUNCTIONS
  // ----------------------

  // Builds a single form to:
  // 1) Pick a template if none selected
  // 2) Optional "Go to created project"
  // 3) Campaign fields (name, artist, date, verticals)
  async function generateTemplateForm(preSelectedTemplate) {
    const form = new Form()
    const templateLibrary = this.templateLibrary
    const preferences = new Preferences('com.KaitlinSalzke.Templates')
    const templateFolder = await templateLibrary.getTemplateFolder()

    // If no template was pre-selected, prompt user to choose one
    if (!preSelectedTemplate) {
      const templateProjects = templateFolder.flattenedProjects.filter(project => {
        let includeOnHold = preferences.readBoolean('includeOnHoldProjects')
        let isOnHold = (includeOnHold && project.status === Project.Status.OnHold)
        let isActive = (project.status === Project.Status.Active)
        return isActive || isOnHold
      })
      form.addField(
        new Form.Field.Option(
          'template',
          'Template',
          templateProjects,
          templateProjects.map(p => p.name),
          null
        )
      )
    }

    // Check if user wants to jump to the new project afterwards
    form.addField(
      new Form.Field.Checkbox(
        'goTo',
        'Go to created project',
        preferences.readBoolean('alwaysGoTo')
      )
    )

    // Additional campaign fields
    form.addField(new Form.Field.String('campaignName', 'Campaign Name', '«Campaign Name»'))
    form.addField(new Form.Field.String('artistName', 'Artist Name', '«Artist Name»'))
    form.addField(new Form.Field.Date('releaseDate', 'Release Date', new Date()))
    form.addField(new Form.Field.MultipleOptions(
      'selectedVerticals',
      'Select Promotional Verticals',
      ['DSP', 'Radio', 'Press', 'DJ'],
      ['DSP', 'Radio', 'Press', 'DJ']
    ))

    return form
  }

  // Replaces placeholders (e.g. «Campaign Name») in task names & notes
  function replacePlaceholders(project, replacements) {
    project.tasks.forEach(task => {
      for (const [placeholder, value] of Object.entries(replacements)) {
        task.name = task.name.replace(placeholder, value)
        if (task.note) {
          task.note = task.note.replace(placeholder, value)
        }
      }
    })
  }

  // Removes tasks that have tags matching any vertical the user did NOT select
  function removeUnselectedVerticalTasks(project, selectedVerticals) {
    const allVerticals = ['DSP', 'Radio', 'Press', 'DJ']
    const verticalsToRemove = allVerticals.filter(v => !selectedVerticals.includes(v))

    for (let i = project.tasks.length - 1; i >= 0; i--) {
      const task = project.tasks[i]
      verticalsToRemove.forEach(vertical => {
        const verticalTag = task.tags.find(tag => tag.name === vertical)
        if (verticalTag) {
          task.remove()
        }
      })
    }
  }

  // Finds $DEFER= / $DUE= lines in the note, calculates offset from releaseDate
  function applyDateOffsets(project, releaseDate) {
    project.tasks.forEach(task => {
      if (task.note) {
        const deferMatch = task.note.match(/\$DEFER=([^\n]+)/)
        if (deferMatch) {
          task.deferDate = calculateOffset(releaseDate, deferMatch[1])
        }
        const dueMatch = task.note.match(/\$DUE=([^\n]+)/)
        if (dueMatch) {
          task.dueDate = calculateOffset(releaseDate, dueMatch[1])
        }
      }
    })
  }

  // Basic offset parser for expressions like "-63d", "+1w", "-2m", "+1y"
  function calculateOffset(referenceDate, offsetString) {
    const match = offsetString.match(/([+-]?)(\d+)([dwmy])/)
    if (!match) return referenceDate

    const sign = (match[1] === '-') ? -1 : 1
    const value = parseInt(match[2]) * sign
    const unit = match[3]

    const newDate = new Date(referenceDate)
    switch (unit) {
      case 'd':
        newDate.setDate(newDate.getDate() + value)
        break
      case 'w':
        newDate.setDate(newDate.getDate() + (value * 7))
        break
      case 'm':
        newDate.setMonth(newDate.getMonth() + value)
        break
      case 'y':
        newDate.setFullYear(newDate.getFullYear() + value)
        break
    }
    return newDate
  }

  return action
})()
