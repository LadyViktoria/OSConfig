
::	Configure Power Settings
powercfg -x -standby-timeout-ac 0
powercfg -x -standby-timeout-dc 30
powercfg -x -hibernate-timeout-ac 0
powercfg -x -hibernate-timeout-dc 60

::	Enable High Performance
::	powercfg /s 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c