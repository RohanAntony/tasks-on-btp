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
  }

  entity TaskHistory : managed {
    key ID          : UUID;
        task        : Association to Tasks;
        field       : String;
        oldValue    : String;
        newValue    : String;
  }
}
