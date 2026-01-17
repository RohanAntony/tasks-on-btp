type TaskStatus : String enum {
  Open        = 'open';
  InProgress  = 'in_progress';
  Review      = 'review';
  Completed   = 'completed';
}

@odata service TasksService {
  entity Tasks {
    key ID          : UUID;
        title       : String;
        description : String;
        dueDate     : Date;
        status      : TaskStatus;
  }
}
