
Public Function doWork(Optional start=Nothing, Optional ignore_failed_puts=1, Optional maximum_work=Nothing, Optional maximum_time=Nothing)
   ' Scan down a tree
   doWork = myObj.doWork(start, ignore_failed_puts, maximum_work, maximum_time)
End Function

Public Function externalQuitRequested()
   ' Check for an external request for a quit
   externalQuitRequested = myObj.externalQuitRequested()
End Function

Public Function isValidSolution(soln, reason)
   ' Return false for a last chance to veto this solution
   ' 
   ' 		If this method returns false then, by default, the children of this packet
   ' 		will be processed normally. If you don't want this to happen then you should
   ' 		also set the status of the packet to INVALID.
   ' 
   ' 		
   isValidSolution = myObj.isValidSolution(soln, reason)
End Function

Public Function newWorker()
   ' Create a new worker of the same type to solve another problem
   newWorker = myObj.newWorker()
End Function