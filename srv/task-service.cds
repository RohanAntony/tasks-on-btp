using { managed } from '@sap/cds/common';

type TaskStatus : String enum {
  Open        = 'open';
  InProgress  = 'in_progress';
  Review      = 'review';
  Completed   = 'completed';
}

@odata service TasksService {
  entity Tasks : managed {
    key ID          : UUID;
        title       : String;
        description : String;
        dueDate     : Date;
        status      : TaskStatus;
        history     : Composition of many TaskHistory on history.task = $self;
        tags        : Composition of many TaskTags on tags.task = $self;
        comments    : Composition of many TaskComments on comments.task = $self;
  }

  entity Tags {
    key ID         : UUID;
        name       : String(50);
        color      : String(7);
        taskCount  : Integer;
  }

  entity TaskTags {
    key task  : Association to Tasks;
    key tag   : Association to Tags;
  }

  entity TaskHistory : managed {
    key ID          : UUID;
        task        : Association to Tasks;
        field       : String;
        oldValue    : String;
        newValue    : String;
  }

  entity TaskComments : managed {
    key ID      : UUID;
        task    : Association to Tasks;
        content : String;
  }
}
